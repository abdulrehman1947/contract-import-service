import express, { Request, Response } from 'express';
import multer from 'multer';
import * as xlsx from 'xlsx';
import mysql from 'mysql2/promise';
import dotenv from 'dotenv';
import cors from 'cors';
import axios from 'axios';

dotenv.config();

const app = express();
// app.use(cors());
app.use(cors({
  origin: '*',
  methods: ['GET', 'POST', 'PUT', 'DELETE', 'OPTIONS'],
  allowedHeaders: ['Content-Type', 'Authorization'],
}));
app.use(express.json());

// Setup DB Pool
const pool = mysql.createPool({
    host: process.env.DB_HOST,
    user: process.env.DB_USER,
    password: process.env.DB_PASSWORD,
    database: process.env.DB_NAME,
    waitForConnections: true,
    connectionLimit: 10
});

// Configure Multer for File Upload
const storage = multer.memoryStorage();
const upload = multer({ storage });

interface ImportLog {
    row: number;
    raison_sociale: string;
    status: 'SUCCESS' | 'WARNING' | 'ERROR' | 'SKIPPED';
    message: string;
}
// Helper to fetch INSEE Data
async function fetchInseeData(siret: string) {
    if (!siret || siret.length !== 14) return null;
    try {
        const response = await axios.get(`https://api.insee.fr/api-sirene/3.11/siret/${siret}`, {
            headers: {
                "X-INSEE-Api-Key-Integration": process.env.INSEE_API_KEY,
                "Accept": "application/json"
            },
            timeout: 5000 
        });
        const est = response.data?.etablissement;
        if (!est) return null;
        const addr = est.adresseEtablissement;
        const street = [
            addr?.numeroVoieEtablissement,
            addr?.indiceRepetitionEtablissement,
            addr?.typeVoieEtablissement,
            addr?.libelleVoieEtablissement
        ].filter(Boolean).join(" ");

        return {
            society: est.uniteLegale?.denominationUniteLegale || null,
            nafCode: est.uniteLegale?.activitePrincipaleUniteLegale || null,
            street: street,
            city: addr?.libelleCommuneEtablissement || null,
            postalCode: addr?.codePostalEtablissement || null
        };
    } catch (error: any) {
        return null;
    }
}
app.post('/api/import-excel', upload.single('file'), async (req: Request, res: Response) => {
    if (!req.file) {
        return res.status(400).json({ error: 'No file uploaded' });
    }

    const logs: ImportLog[] = [];
    const workbook = xlsx.read(req.file.buffer, { type: 'buffer' });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const data: any[] = xlsx.utils.sheet_to_json(sheet);

    const connection = await pool.getConnection();
    // Use provided default employee id from the form-data if present, otherwise fall back to legacy id
    const defaultEmployeeId: number = parseInt(String(req.body?.defaultEmployeeId || ''), 10) ;

    try {
        for (let i = 0; i < data.length; i++) {
            const row = data[i];
            const rowNum = i + 2; // Excel header is row 1
            
            // Extract and clean Excel data
            const rawCompteur = String(row['compteur'] || '').trim();
            const rawType = String(row['type'] || '').trim().toUpperCase();
            const raisonSociale = String(row['raison sociale'] || '').trim();
            const emailClient = String(row['Email Client'] || '').trim();
            const siren = String(row['siren'] ||row['siret'] || '').trim(); // Fallback column if added
            const fournisseur = String(row['fournisseur'] || '').trim();
            const dateDF = String(row['dateDF'] || '').trim();
            const dateFF = String(row['dateFF'] || '').trim();
            const excelDate = String(row['date'] || '').trim();
            const consommation = String(row['consommation'] || '0').replace(',', '.');
            const marge = String(row['marge'] || '0').replace(',', '.');
            const brut = String(row['brut'] || '0').replace(',', '.');
            const duree = parseInt(row['duree']) || 0;
            const vendeur = String(row['vendeur'] || '').trim();
            const signedContractDate = String(row['date'] || '').trim();

            // Extra client info for auto-creation
            const nomClient = String(row['Nom de client'] || '').trim();
            const prenomClient = String(row['Prenom de client'] || '').trim();
            const telClient = String(row['N de tel'] || '').trim();
            const scoreCreditsafe = String(row['Score Creditsafe'] || '').trim();
            const scoreEllipro = String(row['Score Ellipro'] || '').trim();

            const rowSuccessMessages: string[] = [];

            if (!rawCompteur || !rawType) {
                logs.push({ row: rowNum, raison_sociale: raisonSociale, status: 'ERROR', message: 'Missing Meter number or Energy Type' });
                continue;
            }
            console.log(`Processing row ${rowNum}: Compteur=${rawCompteur}, Type=${rawType}, Raison Sociale=${raisonSociale}`);
            try {

                // FETCH INSEE DATA BEFORE TRANSACTION
                const insee = await fetchInseeData(siren);
                await connection.beginTransaction();

                // 1. DUPLICATE CONTRACT CHECK
                // We check this BEFORE reactivating clients or meters to avoid log/db mismatch
                const isElec = rawType.includes('ELEC');
                const contractMeterCol = isElec ? 'pdl' : 'pce_number';
                const [existingCtrs]: any = await connection.execute(
                    `SELECT id, status FROM contracts WHERE ${contractMeterCol} = ? AND status IN ('PARTNER_VERIFIED', 'CANCELLED')`, 
                    [rawCompteur]
                );
                  if (existingCtrs.length > 0) {
                    logs.push({ row: rowNum, raison_sociale: raisonSociale, status: 'SKIPPED', message: `Active contract (${existingCtrs[0].status}) already exists for meter ${rawCompteur}.` });
                    await connection.rollback();
                    continue;
                }
                // 1. FIND CLIENT (Logic: Siren > Email > Society Name)
                 let [clients]: any = await connection.execute(
                    `SELECT id, flag, society, email, first_name, last_name, phone, siren, city, country, postal_code, street 
                     FROM clients 
                     WHERE (siren = ? AND siren != '') OR (email = ? AND email != '') OR (society = ?) OR (raison = ?) LIMIT 1`,
                    [siren || null, emailClient || null, raisonSociale, raisonSociale]
                );

                let client;
                if (clients.length === 0) {
                    // --- CREATE CLIENT IF NOT EXISTS ---
                    // Create Client using INSEE fallbacks
                    const [newClientResult]: any = await connection.execute(
                        `INSERT INTO clients (
                            society, trade_name, raison, score_credit_safe, score_ellipro,
                            last_name, first_name, phone, email, siren,
                            street, city, postal_code, country, naf_code,
                            flag, created_by, created_on, updated_on, type
                        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'France', ?, 'ACTIVE', 'EXCEL_IMPORT', STR_TO_DATE(?, '%d/%m/%Y'), NOW(), 'CLIENT')`,
                        [
                            insee?.society || raisonSociale, 
                            insee?.society || raisonSociale, 
                            raisonSociale, 
                            row['Score Creditsafe'] || '', 
                            row['Score Ellipro'] || '',
                            nomClient, prenomClient, telClient, emailClient, siren,
                            insee?.street || '', insee?.city || '', insee?.postalCode || '', insee?.nafCode || '', signedContractDate
                        ]
                    );
                    
                    const newClientId = newClientResult.insertId;
                     // Create Contact Record (Crucial for DB consistency)
                    await connection.execute(
                        `INSERT INTO client_contact (
                            client_id, first_name, last_name, email, phone, mobile_phone, 
                            job_function, civility, flag, created_by, created_on, updated_on
                        ) VALUES (?, ?, ?, ?, ?, ?, 'Gérant', 'MR', 'ACTIVE', 'EXCEL_IMPORT', STR_TO_DATE(?, '%d/%m/%Y'), NOW())`,
                        [newClientId, prenomClient, nomClient, emailClient, telClient, telClient, signedContractDate]
                    );
                    // Fetch the created client to populate the snapshot correctly
                    let [newClients]: any = await connection.execute(`SELECT * FROM clients WHERE id = ?`, [newClientId]);
                    client = newClients[0];
                    rowSuccessMessages.push(insee ? "Client created (INSEE)" : "Client created");
                } else {
                    client = clients[0];
                     // --- REACTIVATION CHECK FOR CLIENT ---
                    if (client.flag === 'DELETED') {
                        await connection.execute(`UPDATE clients SET flag = 'ACTIVE', updated_on = NOW() WHERE id = ?`, [client.id]);
                        rowSuccessMessages.push("Client reactivated");
                        
                        // Also check if its contact is deleted and reactivate it
                        // await connection.execute(`UPDATE client_contact SET flag = 'ACTIVE', updated_on = NOW() WHERE client_id = ? AND flag = 'DELETED'`, [client.id]);
                    }
                }

                // 2. FIND EMPLOYEE
                let [employees]: any = await connection.execute(
                    `SELECT id FROM employees WHERE TRIM(REPLACE(?, '  ', ' ')) = CONCAT(first_name, ' ', name) 
                     OR TRIM(REPLACE(?, '  ', ' ')) = CONCAT(name, ' ', first_name) LIMIT 1`,
                    [vendeur, vendeur]
                );
                const employeeId = employees.length > 0 ? employees[0].id : defaultEmployeeId;
                console.log(`Matched Employee ID: ${employeeId} for Vendeur: ${vendeur} (default ${defaultEmployeeId})`);

                // 3. CHECK METER EXISTENCE (client_histories)
                // const isElec = rawType.includes('ELEC');
                const meterCol = isElec ? 'pdl' : 'pce';
                
                let [histories]: any = await connection.execute(
                    `SELECT id FROM client_histories WHERE ${meterCol} = ?`, [rawCompteur]
                );

                if (histories.length === 0) {
                    // Create history record if missing
                    await connection.execute(
                        `INSERT INTO client_histories (client_id, ${meterCol}, type, car_in_mwh, current_contract_expiry_date, contract_case, current_supplier_name, flag, created_by, created_on) 
                         VALUES (?, ?, ?, ?, STR_TO_DATE(?, '%d/%m/%Y'), 'SupplierChange', ?, 'ACTIVE', 'EXCEL_IMPORT', NOW())`,
                        [client.id, rawCompteur, isElec ? 'ELECTRICITY' : 'GAS', consommation, dateFF, fournisseur]
                    );
                    // logs.push({ row: rowNum, raison_sociale: raisonSociale, status: 'SUCCESS', message: `Meter ${rawCompteur} added to client history.` });
                    rowSuccessMessages.push("Meter added");
                } else if (histories[0].flag === 'DELETED') {
                    // --- REACTIVATION CHECK FOR METER ---
                    await connection.execute(`UPDATE client_histories SET flag = 'ACTIVE', client_id = ?, updated_on = NOW() WHERE id = ?`, [client.id, histories[0].id]);
                    // logs.push({ row: rowNum, raison_sociale: raisonSociale, status: 'SUCCESS', message: `Deleted Meter ${rawCompteur} reactivated.` });
                    rowSuccessMessages.push("Meter reactivated");
                }

                // 4. CHECK CONTRACT STATUS
                // const contractMeterCol = isElec ? 'pdl' : 'pce_number';
                // let [contracts]: any = await connection.execute(
                //     `SELECT id, status FROM contracts WHERE ${contractMeterCol} = ?`, [rawCompteur]
                // );

                // if (contracts.length > 0) {
                //     const existingStatus = contracts[0].status;
                //     if (['PARTNER_VERIFIED', 'CANCELLED'].includes(existingStatus)) {
                //         logs.push({ row: rowNum, raison_sociale: raisonSociale, status: 'SKIPPED', message: `Meter ${rawCompteur} already has a ${existingStatus} contract.` });
                //         await connection.rollback();
                //         continue;
                //     }
                // }

                // 5. CREATE CLIENT SNAPSHOT
                const [snapResult]: any = await connection.execute(
                    `INSERT INTO contract_client_snapshots (
                        created_by, created_on, flag, society, email, first_name, last_name, phone, siren, client_id,
                        business_city, business_country, business_postalcode, business_street
                    ) VALUES ('EXCEL_IMPORT', NOW(), 'ACTIVE', ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`,
                    [client.society, client.email, client.first_name, client.last_name, client.phone, client.siren, client.id,
                     client.city, client.country, client.postal_code, client.street]
                );
                const snapshotId = snapResult.insertId;

                // 6. CREATE CONTRACT
                await connection.execute(
                    `INSERT INTO contracts (
                        created_by, created_on, flag, status, client_id, employee_id, client_snapshot_id,
                        type, ${contractMeterCol}, car_in_mwh, consumption_inmwh, total_consumption,
                        current_supplier_name, start_date, contract_end_date, current_contract_expiry_date, 
                        signed_contract_date, desired_duration, margin_volt, volt_unit_margin, contract_case, updated_on
                    ) VALUES (
                        'EXCEL_IMPORT', STR_TO_DATE(?, '%d/%m/%Y'), 'ACTIVE', 'PARTNER_VERIFIED', ?, ?, ?, 
                        ?, ?, ?, ?, ?, ?, 
                        STR_TO_DATE(?, '%d/%m/%Y'), STR_TO_DATE(?, '%d/%m/%Y'), STR_TO_DATE(?, '%d/%m/%Y'),
                        STR_TO_DATE(?, '%d/%m/%Y'), ?, ?, ?, 'SupplierChange', NOW()
                    )`,
                    [
                        signedContractDate, client.id, employeeId, snapshotId,
                        isElec ? 'ELECTRICITY' : 'GAS', rawCompteur, consommation, consommation, consommation,
                        fournisseur, dateDF, dateFF, dateFF, excelDate, duree, brut, marge
                    ]
                );

                await connection.commit();
                // logs.push({ row: rowNum, raison_sociale: raisonSociale, status: 'SUCCESS', message: 'Contract created successfully.' });
                logs.push({ row: rowNum, raison_sociale: raisonSociale, status: 'SUCCESS', message: rowSuccessMessages.length > 0 ? rowSuccessMessages.join(", ") + ". Contract created." : "Contract created." });

            } catch (innerErr: any) {
                await connection.rollback();
                logs.push({ row: rowNum, raison_sociale: raisonSociale, status: 'ERROR', message: `Transaction failed: ${innerErr.message}` });
            }
        }

        res.json({ logs });

    } catch (err: any) {
        res.status(500).json({ error: err.message });
    } finally {
        connection.release();
    }
});

const PORT = process.env.PORT || 5000;
app.listen(PORT, () => console.log(`Server running on port ${PORT}`));