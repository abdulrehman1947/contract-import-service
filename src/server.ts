import express, { Request, Response } from 'express';
import multer from 'multer';
import * as xlsx from 'xlsx';
import mysql from 'mysql2/promise';
import dotenv from 'dotenv';
import cors from 'cors';

dotenv.config();

const app = express();
app.use(cors());
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

app.post('/api/import-excel', upload.single('file'), async (req: Request, res: Response) => {
    if (!req.file) {
        return res.status(400).json({ error: 'No file uploaded' });
    }

    const logs: ImportLog[] = [];
    const workbook = xlsx.read(req.file.buffer, { type: 'buffer' });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const data: any[] = xlsx.utils.sheet_to_json(sheet);

    const connection = await pool.getConnection();

    try {
        for (let i = 0; i < data.length; i++) {
            const row = data[i];
            const rowNum = i + 2; // Excel header is row 1
            
            // Extract and clean Excel data
            const rawCompteur = String(row['compteur'] || '').trim();
            const rawType = String(row['type'] || '').trim().toUpperCase();
            const raisonSociale = String(row['raison sociale'] || '').trim();
            const emailClient = String(row['Email Client'] || '').trim();
            const siren = String(row['siren'] || '').trim(); // Fallback column if added
            const fournisseur = String(row['fournisseur'] || '').trim();
            const dateDF = String(row['dateDF'] || '').trim();
            const dateFF = String(row['dateFF'] || '').trim();
            const excelDate = String(row['date'] || '').trim();
            const consommation = String(row['consommation'] || '0').replace(',', '.');
            const marge = String(row['marge'] || '0').replace(',', '.');
            const brut = String(row['brut'] || '0').replace(',', '.');
            const duree = parseInt(row['duree']) || 0;
            const vendeur = String(row['vendeur'] || '').trim();

            // Extra client info for auto-creation
            const nomClient = String(row['Nom de client'] || '').trim();
            const prenomClient = String(row['Prenom de client'] || '').trim();
            const telClient = String(row['N de tel'] || '').trim();
            const scoreCreditsafe = String(row['Score Creditsafe'] || '').trim();
            const scoreEllipro = String(row['Score Ellipro'] || '').trim();

            if (!rawCompteur || !rawType) {
                logs.push({ row: rowNum, raison_sociale: raisonSociale, status: 'ERROR', message: 'Missing Meter number or Energy Type' });
                continue;
            }
            console.log(`Processing row ${rowNum}: Compteur=${rawCompteur}, Type=${rawType}, Raison Sociale=${raisonSociale}`);
            try {
                await connection.beginTransaction();

                // 1. FIND CLIENT (Logic: Siren > Email > Society Name)
                let [clients]: any = await connection.execute(
                    `SELECT id, society, email, first_name, last_name, phone, siren, city, country, postal_code, street FROM clients 
                     WHERE (siren = ? AND siren != '') OR (email = ? AND email != '') OR (society = ?) OR (raison = ?) LIMIT 1`,
                    [siren || null, emailClient || null, raisonSociale, raisonSociale]
                );

                let client;
                if (clients.length === 0) {
                    // --- CREATE CLIENT IF NOT EXISTS ---
                    const [newClientResult]: any = await connection.execute(
                        `INSERT INTO clients (
                            society, trade_name, raison, score_credit_safe, score_ellipro,
                            last_name, first_name, phone, email, siren,
                            flag, created_by, created_on
                        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'ACTIVE', 'EXCEL_IMPORT', NOW())`,
                        [
                            raisonSociale, raisonSociale, raisonSociale, scoreCreditsafe, scoreEllipro,
                            nomClient, prenomClient, telClient, emailClient, siren
                        ]
                    );
                    
                    const newClientId = newClientResult.insertId;
                    
                    // Fetch the created client to populate the snapshot correctly
                    let [newClients]: any = await connection.execute(`SELECT * FROM clients WHERE id = ?`, [newClientId]);
                    client = newClients[0];
                    logs.push({ row: rowNum, raison_sociale: raisonSociale, status: 'SUCCESS', message: 'New client created automatically.' });
                } else {
                    client = clients[0];
                }

                // 2. FIND EMPLOYEE
                let [employees]: any = await connection.execute(
                    `SELECT id FROM employees WHERE TRIM(REPLACE(?, '  ', ' ')) = CONCAT(first_name, ' ', name) 
                     OR TRIM(REPLACE(?, '  ', ' ')) = CONCAT(name, ' ', first_name) LIMIT 1`,
                    [vendeur, vendeur]
                );
                const employeeId = employees.length > 0 ? employees[0].id : 184443;
                console.log(`Matched Employee ID: ${employeeId} for Vendeur: ${vendeur}`);

                // 3. CHECK METER EXISTENCE (client_histories)
                const isElec = rawType.includes('ELEC');
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
                    logs.push({ row: rowNum, raison_sociale: raisonSociale, status: 'SUCCESS', message: `Meter ${rawCompteur} added to client history.` });
                }

                // 4. CHECK CONTRACT STATUS
                const contractMeterCol = isElec ? 'pdl' : 'pce_number';
                let [contracts]: any = await connection.execute(
                    `SELECT id, status FROM contracts WHERE ${contractMeterCol} = ?`, [rawCompteur]
                );

                if (contracts.length > 0) {
                    const existingStatus = contracts[0].status;
                    if (['PARTNER_VERIFIED', 'CANCELLED'].includes(existingStatus)) {
                        logs.push({ row: rowNum, raison_sociale: raisonSociale, status: 'SKIPPED', message: `Meter ${rawCompteur} already has a ${existingStatus} contract.` });
                        await connection.rollback();
                        continue;
                    }
                }

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
                        signed_contract_date, desired_duration, margin_volt, volt_unit_margin, contract_case
                    ) VALUES (
                        'EXCEL_IMPORT', NOW(), 'ACTIVE', 'PARTNER_VERIFIED', ?, ?, ?, 
                        ?, ?, ?, ?, ?, ?, 
                        STR_TO_DATE(?, '%d/%m/%Y'), STR_TO_DATE(?, '%d/%m/%Y'), STR_TO_DATE(?, '%d/%m/%Y'),
                        STR_TO_DATE(?, '%d/%m/%Y'), ?, ?, ?, 'SupplierChange'
                    )`,
                    [
                        client.id, employeeId, snapshotId,
                        isElec ? 'ELECTRICITY' : 'GAS', rawCompteur, consommation, consommation, consommation,
                        fournisseur, dateDF, dateFF, dateFF, excelDate, duree, brut, marge
                    ]
                );

                await connection.commit();
                logs.push({ row: rowNum, raison_sociale: raisonSociale, status: 'SUCCESS', message: 'Contract created successfully.' });

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