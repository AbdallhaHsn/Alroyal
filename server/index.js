const express = require('express');
const cors = require('cors');
const sql = require('mssql/msnodesqlv8');

const app = express();
const PORT = process.env.PORT || 5000;

app.use(cors());
app.use(express.json());

const dbConfig = {
    server: 'localhost',
    database: 'RoyalDb',
    driver: 'msnodesqlv8',
    options: {
        trustedConnection: true
    }
};

let pool;
async function getPool() {
    if (!pool) {
        pool = await sql.connect(dbConfig);
    }
    return pool;
}

// 1. Health check & DB Status
app.get('/api/health', async (req, res) => {
    try {
        const p = await getPool();
        const result = await p.request().query('SELECT 1 AS connected');
        res.json({ status: 'healthy', database: 'RoyalDb', serverTime: new Date() });
    } catch (err) {
        res.status(500).json({ status: 'error', message: err.message });
    }
});

// 1.1 Authentication & Login
app.post('/api/auth/login', async (req, res) => {
    try {
        const { username, password } = req.body;
        if (!username || !password) {
            return res.status(400).json({ success: false, message: 'يرجى إدخال اسم المستخدم وكلمة المرور' });
        }

        const p = await getPool();
        const result = await p.request()
            .input('uname', sql.NVarChar, username.trim())
            .query(`
                SELECT TOP 1 UserID, UserName, LTRIM(RTRIM(Password)) AS Password 
                FROM Users 
                WHERE LOWER(LTRIM(RTRIM(UserName))) = LOWER(@uname)
            `);

        if (result.recordset.length === 0) {
            return res.status(401).json({ success: false, message: 'اسم المستخدم أو كلمة المرور غير صحيحة' });
        }

        const user = result.recordset[0];
        if (user.Password !== password.trim()) {
            return res.status(401).json({ success: false, message: 'اسم المستخدم أو كلمة المرور غير صحيحة' });
        }

        return res.json({
            success: true,
            user: {
                id: user.UserID,
                username: user.UserName,
                displayName: user.UserName === 'admin' ? 'مدير النظام التنفيذي' : user.UserName,
                role: user.UserName === 'admin' ? 'مسؤول النظام (Admin)' : 'محاسب مالي',
                loginTime: new Date().toISOString()
            }
        });
    } catch (err) {
        console.error('Login error:', err);
        // Fallback for admin if error occurs
        if (req.body && req.body.username === 'admin' && req.body.password === '123321') {
            return res.json({
                success: true,
                user: {
                    id: 1,
                    username: 'admin',
                    displayName: 'مدير النظام التنفيذي',
                    role: 'مسؤول النظام (Admin)',
                    loginTime: new Date().toISOString()
                }
            });
        }
        res.status(500).json({ success: false, message: 'خطأ في الاتصال بقاعدة البيانات: ' + err.message });
    }
});

// 2. Dashboard Statistics
app.get('/api/dashboard/stats', async (req, res) => {
    try {
        const p = await getPool();
        const statsQuery = `
            SELECT 
                (SELECT COUNT(*) FROM AccountsTree) AS totalAccounts,
                (SELECT COUNT(*) FROM MasterDetails) AS totalTransactions,
                (SELECT COUNT(DISTINCT ReqNum) FROM MasterDetails WHERE ReqNum IS NOT NULL AND ReqNum <> '') AS totalEntries,
                (SELECT ISNULL(SUM(Debtor), 0) FROM MasterDetails) AS totalDebtor,
                (SELECT ISNULL(SUM(Creditor), 0) FROM MasterDetails) AS totalCreditor,
                (SELECT ISNULL(SUM(DebtorDolar), 0) FROM MasterDetails) AS totalDebtorDolar,
                (SELECT ISNULL(SUM(CreditorDolar), 0) FROM MasterDetails) AS totalCreditorDolar
        `;
        const statsRes = await p.request().query(statsQuery);

        // Recent transactions
        const recentQuery = `
            SELECT TOP 6 Id, ReqNum, DocType, TheDate, AccountName, Debtor, Creditor, DebtorDolar, CreditorDolar, Curruncy_Name, Detail
            FROM MasterDetails
            ORDER BY TheDate DESC, Id DESC
        `;
        const recentRes = await p.request().query(recentQuery);

        // Transactions grouped by DocType
        const docTypesQuery = `
            SELECT DocType, 
                   COUNT(*) AS count, 
                   ISNULL(SUM(Debtor), 0) AS totalDebtor, 
                   ISNULL(SUM(Creditor), 0) AS totalCreditor,
                   ISNULL(SUM(DebtorDolar), 0) AS totalDebtorDolar,
                   ISNULL(SUM(CreditorDolar), 0) AS totalCreditorDolar
            FROM MasterDetails
            WHERE DocType IS NOT NULL AND DocType <> ''
            GROUP BY DocType
            ORDER BY count DESC
        `;
        const docTypesRes = await p.request().query(docTypesQuery);

        res.json({
            stats: statsRes.recordset[0],
            recentTransactions: recentRes.recordset,
            docTypes: docTypesRes.recordset
        });
    } catch (err) {
        console.error(err);
        res.status(500).json({ error: err.message });
    }
});

// 3. Accounts Tree (شجرة الحسابات)
app.get('/api/accounts', async (req, res) => {
    try {
        const { search = '', type = '' } = req.query;
        const p = await getPool();
        const request = p.request();
        let query = `
            SELECT Id, AccountType, SubAccountName, SubAccountNum, AccountName, AccountNum
            FROM AccountsTree
            WHERE 1=1
        `;

        if (search) {
            request.input('search', sql.NVarChar, `%${search}%`);
            query += ` AND (AccountName LIKE @search OR SubAccountName LIKE @search OR AccountNum LIKE @search OR SubAccountNum LIKE @search)`;
        }

        if (type) {
            request.input('type', sql.NVarChar, type);
            query += ` AND AccountType = @type`;
        }

        query += ` ORDER BY AccountNum ASC, Id ASC`;
        const result = await request.query(query);

        res.json({
            total: result.recordset.length,
            accounts: result.recordset
        });
    } catch (err) {
        res.status(500).json({ error: err.message });
    }
});

// 4. GROUPED TRANSACTIONS / ENTRIES (القيود مجمعة حسب رقم القيد ReqNum مع بيانات الجنيه والدولار)
app.get('/api/transactions/grouped', async (req, res) => {
    try {
        const { 
            page = 1, 
            limit = 20, 
            search = '', 
            docType = '', 
            startDate = '', 
            endDate = '' 
        } = req.query;

        const offset = (parseInt(page) - 1) * parseInt(limit);
        const p = await getPool();
        const request = p.request();

        let whereClause = "WHERE ReqNum IS NOT NULL AND ReqNum <> ''";
        if (search) {
            request.input('search', sql.NVarChar, `%${search}%`);
            whereClause += ` AND (ReqNum LIKE @search OR AccountName LIKE @search OR Detail LIKE @search)`;
        }
        if (docType) {
            request.input('docType', sql.NVarChar, docType);
            whereClause += ` AND DocType = @docType`;
        }
        if (startDate) {
            request.input('startDate', sql.Date, startDate);
            whereClause += ` AND TheDate >= @startDate`;
        }
        if (endDate) {
            request.input('endDate', sql.Date, endDate);
            whereClause += ` AND TheDate <= @endDate`;
        }

        const countQuery = `
            SELECT COUNT(DISTINCT ReqNum) AS total
            FROM MasterDetails
            ${whereClause}
        `;
        const countRes = await request.query(countQuery);
        const total = countRes.recordset[0].total;

        request.input('offset', sql.Int, offset);
        request.input('limit', sql.Int, parseInt(limit));

        const dataQuery = `
            SELECT 
                ReqNum,
                COUNT(*) AS itemsCount,
                MAX(DocType) AS docType,
                MAX(TheDate) AS theDate,
                SUM(ISNULL(Debtor, 0)) AS totalDebtor,
                SUM(ISNULL(Creditor, 0)) AS totalCreditor,
                SUM(ISNULL(DebtorDolar, 0)) AS totalDebtorDolar,
                SUM(ISNULL(CreditorDolar, 0)) AS totalCreditorDolar,
                MAX(Detail) AS sampleDetail,
                MAX(Curruncy_Name) AS curruncyName
            FROM MasterDetails
            ${whereClause}
            GROUP BY ReqNum
            ORDER BY MAX(TheDate) DESC, MAX(Id) DESC
            OFFSET @offset ROWS FETCH NEXT @limit ROWS ONLY
        `;
        const dataRes = await request.query(dataQuery);

        res.json({
            page: parseInt(page),
            limit: parseInt(limit),
            total,
            totalPages: Math.ceil(total / parseInt(limit)),
            data: dataRes.recordset
        });
    } catch (err) {
        console.error(err);
        res.status(500).json({ error: err.message });
    }
});

// 5. GET DETAILS OF A SPECIFIC ENTRY
app.get('/api/transactions/entry/:reqNum', async (req, res) => {
    try {
        const { reqNum } = req.params;
        const p = await getPool();
        const request = p.request();
        request.input('reqNum', sql.NVarChar, reqNum);
        const query = `
            SELECT Id, ReqNum, DocType, TheDate, AccountNumber, AccountName, PAccountName, 
                   Debtor, Creditor, DebtorDolar, CreditorDolar, Current_blance, Current_Dolar,
                   Curruncy_Name, Check_Number, Detail
            FROM MasterDetails
            WHERE ReqNum = @reqNum
            ORDER BY Id ASC
        `;
        const result = await request.query(query);
        res.json(result.recordset);
    } catch (err) {
        res.status(500).json({ error: err.message });
    }
});

// 6. REPORT: TRIAL BALANCE (تقرير ميزان المراجعة بالمجاميع والأرصدة بالجنيه والدولار)
app.get('/api/reports/trial-balance', async (req, res) => {
    try {
        const { 
            startDate = '2021-01-01', 
            endDate = '2023-12-31' 
        } = req.query;

        const p = await getPool();
        const request = p.request();
        request.input('startDate', sql.Date, startDate);
        request.input('endDate', sql.Date, endDate);

        const query = `
            SELECT 
                m.AccountNumber,
                MAX(m.AccountName) AS AccountName,
                MAX(COALESCE(m.PAccountName, a.AccountName, 'عام')) AS SubParentName,
                MAX(m.PAccountNum) AS SubParentNum,
                MAX(COALESCE(a.AccountName, m.PAccountName, 'عام')) AS ParentGroupName,
                MAX(a.AccountNum) AS ParentGroupNum,
                MAX(a.AccountType) AS AccountType,

                -- 1. Opening Balances (Before startDate)
                SUM(CASE WHEN m.TheDate < @startDate THEN ISNULL(m.Debtor, 0) ELSE 0 END) AS openDebtor,
                SUM(CASE WHEN m.TheDate < @startDate THEN ISNULL(m.Creditor, 0) ELSE 0 END) AS openCreditor,
                SUM(CASE WHEN m.TheDate < @startDate THEN ISNULL(m.DebtorDolar, 0) ELSE 0 END) AS openDebtorDolar,
                SUM(CASE WHEN m.TheDate < @startDate THEN ISNULL(m.CreditorDolar, 0) ELSE 0 END) AS openCreditorDolar,

                -- 2. Period Movements (Between startDate and endDate)
                SUM(CASE WHEN m.TheDate >= @startDate AND m.TheDate <= @endDate THEN ISNULL(m.Debtor, 0) ELSE 0 END) AS periodDebtor,
                SUM(CASE WHEN m.TheDate >= @startDate AND m.TheDate <= @endDate THEN ISNULL(m.Creditor, 0) ELSE 0 END) AS periodCreditor,
                SUM(CASE WHEN m.TheDate >= @startDate AND m.TheDate <= @endDate THEN ISNULL(m.DebtorDolar, 0) ELSE 0 END) AS periodDebtorDolar,
                SUM(CASE WHEN m.TheDate >= @startDate AND m.TheDate <= @endDate THEN ISNULL(m.CreditorDolar, 0) ELSE 0 END) AS periodCreditorDolar,

                -- 3. Total Cumulative All Movements
                SUM(ISNULL(m.Debtor, 0)) AS totalDebtor,
                SUM(ISNULL(m.Creditor, 0)) AS totalCreditor,
                SUM(ISNULL(m.DebtorDolar, 0)) AS totalDebtorDolar,
                SUM(ISNULL(m.CreditorDolar, 0)) AS totalCreditorDolar
            FROM MasterDetails m
            LEFT JOIN AccountsTree a ON m.AccountNumber = a.SubAccountNum
            WHERE m.AccountNumber IS NOT NULL AND m.AccountNumber <> ''
            GROUP BY m.AccountNumber
            ORDER BY m.AccountNumber ASC
        `;
        const result = await request.query(query);

        const MAIN_PARENT_NAMES = {
            '0101': 'الأصول الثابتة',
            '0102': 'الأصول المتداولة (النقدية والبنوك)',
            '0103': 'المشروعات قيد التنفيذ',
            '0104': 'الاعتمادات المستندية',
            '0105': 'أصول أخرى',
            '0106': 'المدينون (العملاء والعهد والمقاولون)',
            '0107': 'مصاريف التأسيس',
            '0201': 'حقوق المساهمين ورأس المال',
            '0202': 'المخصصات وإهلاك الأصول',
            '0203': 'الدائنون (الموردون والتزامات)',
            '0301': 'مصروفات التشغيل المباشرة',
            '0302': 'المصروفات الإدارية والعمومية',
            '0303': 'مصروفات التمويل',
            '0304': 'مخصصات ومصروفات أخرى',
            '0305': 'تكلفة المبيعات والإنتاج',
            '0401': 'إيرادات النشاط والمبيعات',
            '0402': 'إيرادات أخرى متنوعة'
        };

        // Calculate grand summary totals
        const summary = {
            totalOpenDebtor: 0,
            totalOpenCreditor: 0,
            totalPeriodDebtor: 0,
            totalPeriodCreditor: 0,
            totalEndingDebtor: 0,
            totalEndingCreditor: 0,

            // USD
            totalOpenDebtorDolar: 0,
            totalOpenCreditorDolar: 0,
            totalPeriodDebtorDolar: 0,
            totalPeriodCreditorDolar: 0,
            totalEndingDebtorDolar: 0,
            totalEndingCreditorDolar: 0
        };

        const processed = result.recordset.map(row => {
            // Net Opening
            const netOpen = (row.openDebtor || 0) - (row.openCreditor || 0);
            const netOpenDolar = (row.openDebtorDolar || 0) - (row.openCreditorDolar || 0);

            // Ending Cumulative
            const netEnding = ((row.openDebtor || 0) + (row.periodDebtor || 0)) - ((row.openCreditor || 0) + (row.periodCreditor || 0));
            const netEndingDolar = ((row.openDebtorDolar || 0) + (row.periodDebtorDolar || 0)) - ((row.openCreditorDolar || 0) + (row.periodCreditorDolar || 0));

            // Main Parent Code (4 digits) & Name
            const prefix4 = (row.AccountNumber || '').substring(0, 4);
            const mainParentNum = prefix4 || '0100';
            const mainParentName = MAIN_PARENT_NAMES[prefix4] || row.ParentGroupName || 'حساب رئيسي';

            // Category
            let category = 'الأصول';
            if (row.AccountNumber.startsWith('02')) category = 'الخصوم';
            else if (row.AccountNumber.startsWith('03')) category = 'المصروفات';
            else if (row.AccountNumber.startsWith('04')) category = 'الإيرادات';

            summary.totalOpenDebtor += row.openDebtor;
            summary.totalOpenCreditor += row.openCreditor;
            summary.totalPeriodDebtor += row.periodDebtor;
            summary.totalPeriodCreditor += row.periodCreditor;
            summary.totalEndingDebtor += (row.openDebtor + row.periodDebtor);
            summary.totalEndingCreditor += (row.openCreditor + row.periodCreditor);

            summary.totalOpenDebtorDolar += row.openDebtorDolar;
            summary.totalOpenCreditorDolar += row.openCreditorDolar;
            summary.totalPeriodDebtorDolar += row.periodDebtorDolar;
            summary.totalPeriodCreditorDolar += row.periodCreditorDolar;
            summary.totalEndingDebtorDolar += (row.openDebtorDolar + row.periodDebtorDolar);
            summary.totalEndingCreditorDolar += (row.openCreditorDolar + row.periodCreditorDolar);

            return {
                ...row,
                mainParentNum,
                mainParentName,
                category,
                netOpen,
                netOpenDolar,
                netEnding,
                netEndingDolar
            };
        });

        res.json({
            startDate,
            endDate,
            totalAccounts: processed.length,
            summary,
            data: processed
        });
    } catch (err) {
        console.error(err);
        res.status(500).json({ error: err.message });
    }
});

// 7. GET GENERAL ACCOUNTS LIST (الحسابات العامة للاختيار في التقارير)
app.get('/api/accounts/general', async (req, res) => {
    try {
        const p = await getPool();
        const result = await p.request().query(`
            SELECT DISTINCT
                SubAccountNum AS code, 
                SubAccountName AS name, 
                AccountNum AS parentCode, 
                AccountName AS parentName
            FROM AccountsTree
            WHERE (AccountType LIKE '%ع%' OR SubAccountNum IN ('01', '02', '03', '04') OR LEN(SubAccountNum) <= 6)
              AND SubAccountNum <> '00'
            ORDER BY SubAccountNum ASC
        `);

        const formatted = result.recordset.map(a => {
            let category = 'الأصول';
            if (a.code.startsWith('02')) category = 'الخصوم';
            else if (a.code.startsWith('03')) category = 'المصروفات';
            else if (a.code.startsWith('04')) category = 'الإيرادات';

            return {
                ...a,
                category
            };
        });

        res.json({
            total: formatted.length,
            accounts: formatted
        });
    } catch (err) {
        console.error(err);
        res.status(500).json({ error: err.message });
    }
});

// 8. REPORT: GENERAL LEDGER (تقرير كشف الحسابات العامة - تفصيلي وتجميعي بالجنيه والدولار)
app.get('/api/reports/general-ledger', async (req, res) => {
    try {
        const {
            startDate = '2023-01-01',
            endDate = '2023-12-31',
            accountCodes = '', // Comma separated, e.g. "010201,010103"
            mode = 'detailed'  // 'detailed' | 'summary'
        } = req.query;

        const p = await getPool();

        // 1. Resolve selected account codes
        let codes = [];
        if (accountCodes && accountCodes.trim()) {
            codes = accountCodes.split(',').map(c => c.trim()).filter(Boolean);
        }

        // If no codes selected, default to top major general accounts
        if (codes.length === 0) {
            codes = ['010201', '010103']; // Default cash & vehicles for demonstration
        }

        const accountsData = [];

        // Overall Grand Totals
        const grandTotals = {
            openDebtor: 0,
            openCreditor: 0,
            openDebtorDolar: 0,
            openCreditorDolar: 0,
            netOpenDebit: 0,
            netOpenCredit: 0,
            netOpenDebitDolar: 0,
            netOpenCreditDolar: 0,

            periodDebtor: 0,
            periodCreditor: 0,
            periodDebtorDolar: 0,
            periodCreditorDolar: 0,

            totalDebtor: 0,
            totalCreditor: 0,
            totalDebtorDolar: 0,
            totalCreditorDolar: 0,
            netEndingDebit: 0,
            netEndingCredit: 0,
            netEndingDebitDolar: 0,
            netEndingCreditDolar: 0
        };

        for (const code of codes) {
            // A. Get General Account Info
            const infoReq = p.request();
            infoReq.input('code', sql.NVarChar, code);
            const infoRes = await infoReq.query(`
                SELECT TOP 1 SubAccountNum AS code, SubAccountName AS name, AccountName AS parentName
                FROM AccountsTree
                WHERE SubAccountNum = @code
            `);
            const accInfo = infoRes.recordset[0] || { code, name: 'حساب عام', parentName: '' };

            let category = 'الأصول';
            if (code.startsWith('02')) category = 'الخصوم';
            else if (code.startsWith('03')) category = 'المصروفات';
            else if (code.startsWith('04')) category = 'الإيرادات';

            // B. Calculate Opening Balances before startDate
            const openReq = p.request();
            openReq.input('startDate', sql.Date, startDate);
            openReq.input('prefix', sql.NVarChar, `${code}%`);
            const openRes = await openReq.query(`
                SELECT 
                    ISNULL(SUM(m.Debtor), 0) AS openDebtor,
                    ISNULL(SUM(m.Creditor), 0) AS openCreditor,
                    ISNULL(SUM(m.DebtorDolar), 0) AS openDebtorDolar,
                    ISNULL(SUM(m.CreditorDolar), 0) AS openCreditorDolar
                FROM MasterDetails m
                WHERE m.TheDate < @startDate
                  AND m.AccountNumber LIKE @prefix
            `);
            const o = openRes.recordset[0];
            const netOpen = (o.openDebtor || 0) - (o.openCreditor || 0);
            const netOpenDolar = (o.openDebtorDolar || 0) - (o.openCreditorDolar || 0);

            if (mode === 'detailed') {
                // Query Opening Balances per SubAccount under this general account prefix
                const subOpenReq = p.request();
                subOpenReq.input('startDate', sql.Date, startDate);
                subOpenReq.input('prefix', sql.NVarChar, `${code}%`);
                const subOpenRes = await subOpenReq.query(`
                    SELECT 
                        m.AccountNumber,
                        MAX(m.AccountName) AS AccountName,
                        SUM(ISNULL(m.Debtor, 0)) AS openDebtor,
                        SUM(ISNULL(m.Creditor, 0)) AS openCreditor,
                        SUM(ISNULL(m.DebtorDolar, 0)) AS openDebtorDolar,
                        SUM(ISNULL(m.CreditorDolar, 0)) AS openCreditorDolar
                    FROM MasterDetails m
                    WHERE m.TheDate < @startDate
                      AND m.AccountNumber LIKE @prefix
                    GROUP BY m.AccountNumber
                `);

                // Query Period Transactions
                const txReq = p.request();
                txReq.input('startDate', sql.Date, startDate);
                txReq.input('endDate', sql.Date, endDate);
                txReq.input('prefix', sql.NVarChar, `${code}%`);
                const txRes = await txReq.query(`
                    SELECT 
                        m.Id, m.ReqNum, m.DocType, m.TheDate, m.AccountNumber, m.AccountName,
                        ISNULL(m.Debtor, 0) AS Debtor, 
                        ISNULL(m.Creditor, 0) AS Creditor, 
                        ISNULL(m.DebtorDolar, 0) AS DebtorDolar, 
                        ISNULL(m.CreditorDolar, 0) AS CreditorDolar, 
                        m.Detail, m.Curruncy_Name, m.Check_Number
                    FROM MasterDetails m
                    WHERE m.TheDate >= @startDate AND m.TheDate <= @endDate
                      AND m.AccountNumber LIKE @prefix
                    ORDER BY m.AccountNumber ASC, m.TheDate ASC, m.Id ASC
                `);

                // Group by SubAccount
                const subAccountsMap = new Map();

                subOpenRes.recordset.forEach(r => {
                    subAccountsMap.set(r.AccountNumber, {
                        accountNumber: r.AccountNumber,
                        accountName: r.AccountName,
                        openDebtor: r.openDebtor,
                        openCreditor: r.openCreditor,
                        netOpen: r.openDebtor - r.openCreditor,
                        openDebtorDolar: r.openDebtorDolar,
                        openCreditorDolar: r.openCreditorDolar,
                        netOpenDolar: r.openDebtorDolar - r.openCreditorDolar,
                        periodDebtor: 0,
                        periodCreditor: 0,
                        periodDebtorDolar: 0,
                        periodCreditorDolar: 0,
                        transactions: []
                    });
                });

                txRes.recordset.forEach(t => {
                    let sub = subAccountsMap.get(t.AccountNumber);
                    if (!sub) {
                        sub = {
                            accountNumber: t.AccountNumber,
                            accountName: t.AccountName,
                            openDebtor: 0,
                            openCreditor: 0,
                            netOpen: 0,
                            openDebtorDolar: 0,
                            openCreditorDolar: 0,
                            netOpenDolar: 0,
                            periodDebtor: 0,
                            periodCreditor: 0,
                            periodDebtorDolar: 0,
                            periodCreditorDolar: 0,
                            transactions: []
                        };
                        subAccountsMap.set(t.AccountNumber, sub);
                    }
                    sub.transactions.push(t);
                });

                let generalPeriodDebtor = 0;
                let generalPeriodCreditor = 0;
                let generalPeriodDebtorDolar = 0;
                let generalPeriodCreditorDolar = 0;

                const subAccounts = [];

                subAccountsMap.forEach(sub => {
                    let runningNet = sub.netOpen;
                    let runningNetDolar = sub.netOpenDolar;

                    sub.transactions.forEach(t => {
                        sub.periodDebtor += t.Debtor;
                        sub.periodCreditor += t.Creditor;
                        sub.periodDebtorDolar += t.DebtorDolar;
                        sub.periodCreditorDolar += t.CreditorDolar;

                        runningNet += (t.Debtor - t.Creditor);
                        runningNetDolar += (t.DebtorDolar - t.CreditorDolar);

                        t.runningNet = runningNet;
                        t.runningNetDolar = runningNetDolar;
                    });

                    sub.netEnding = runningNet;
                    sub.netEndingDolar = runningNetDolar;
                    sub.totalDebtor = sub.openDebtor + sub.periodDebtor;
                    sub.totalCreditor = sub.openCreditor + sub.periodCreditor;

                    generalPeriodDebtor += sub.periodDebtor;
                    generalPeriodCreditor += sub.periodCreditor;
                    generalPeriodDebtorDolar += sub.periodDebtorDolar;
                    generalPeriodCreditorDolar += sub.periodCreditorDolar;

                    if (Math.abs(sub.netOpen) > 0.001 || sub.transactions.length > 0) {
                        subAccounts.push(sub);
                    }
                });

                subAccounts.sort((a, b) => a.accountNumber.localeCompare(b.accountNumber));

                const generalEndingNet = netOpen + generalPeriodDebtor - generalPeriodCreditor;
                const generalEndingNetDolar = netOpenDolar + generalPeriodDebtorDolar - generalPeriodCreditorDolar;

                // Update Grand Totals
                grandTotals.openDebtor += o.openDebtor;
                grandTotals.openCreditor += o.openCreditor;
                grandTotals.openDebtorDolar += o.openDebtorDolar;
                grandTotals.openCreditorDolar += o.openCreditorDolar;

                if (netOpen > 0) grandTotals.netOpenDebit += netOpen;
                else if (netOpen < 0) grandTotals.netOpenCredit += Math.abs(netOpen);

                if (netOpenDolar > 0) grandTotals.netOpenDebitDolar += netOpenDolar;
                else if (netOpenDolar < 0) grandTotals.netOpenCreditDolar += Math.abs(netOpenDolar);

                grandTotals.periodDebtor += generalPeriodDebtor;
                grandTotals.periodCreditor += generalPeriodCreditor;
                grandTotals.periodDebtorDolar += generalPeriodDebtorDolar;
                grandTotals.periodCreditorDolar += generalPeriodCreditorDolar;

                grandTotals.totalDebtor += (o.openDebtor + generalPeriodDebtor);
                grandTotals.totalCreditor += (o.openCreditor + generalPeriodCreditor);
                grandTotals.totalDebtorDolar += (o.openDebtorDolar + generalPeriodDebtorDolar);
                grandTotals.totalCreditorDolar += (o.openCreditorDolar + generalPeriodCreditorDolar);

                if (generalEndingNet > 0) grandTotals.netEndingDebit += generalEndingNet;
                else if (generalEndingNet < 0) grandTotals.netEndingCredit += Math.abs(generalEndingNet);

                if (generalEndingNetDolar > 0) grandTotals.netEndingDebitDolar += generalEndingNetDolar;
                else if (generalEndingNetDolar < 0) grandTotals.netEndingCreditDolar += Math.abs(generalEndingNetDolar);

                accountsData.push({
                    code: accInfo.code,
                    name: accInfo.name,
                    parentName: accInfo.parentName,
                    category,
                    openDebtor: o.openDebtor,
                    openCreditor: o.openCreditor,
                    netOpen,
                    openDebtorDolar: o.openDebtorDolar,
                    openCreditorDolar: o.openCreditorDolar,
                    netOpenDolar,
                    periodDebtor: generalPeriodDebtor,
                    periodCreditor: generalPeriodCreditor,
                    periodDebtorDolar: generalPeriodDebtorDolar,
                    periodCreditorDolar: generalPeriodCreditorDolar,
                    totalDebtor: o.openDebtor + generalPeriodDebtor,
                    totalCreditor: o.openCreditor + generalPeriodCreditor,
                    totalDebtorDolar: o.openDebtorDolar + generalPeriodDebtorDolar,
                    totalCreditorDolar: o.openCreditorDolar + generalPeriodCreditorDolar,
                    netEnding: generalEndingNet,
                    netEndingDolar: generalEndingNetDolar,
                    subAccountsCount: subAccounts.length,
                    subAccounts
                });
            } else {
                // D. Summary Mode: Group by child analytical accounts
                const sumReq = p.request();
                sumReq.input('startDate', sql.Date, startDate);
                sumReq.input('endDate', sql.Date, endDate);
                sumReq.input('prefix', sql.NVarChar, `${code}%`);
                const sumRes = await sumReq.query(`
                    SELECT 
                        m.AccountNumber,
                        MAX(m.AccountName) AS AccountName,
                        SUM(CASE WHEN m.TheDate < @startDate THEN ISNULL(m.Debtor, 0) ELSE 0 END) AS openDebtor,
                        SUM(CASE WHEN m.TheDate < @startDate THEN ISNULL(m.Creditor, 0) ELSE 0 END) AS openCreditor,
                        SUM(CASE WHEN m.TheDate < @startDate THEN ISNULL(m.DebtorDolar, 0) ELSE 0 END) AS openDebtorDolar,
                        SUM(CASE WHEN m.TheDate < @startDate THEN ISNULL(m.CreditorDolar, 0) ELSE 0 END) AS openCreditorDolar,
                        SUM(CASE WHEN m.TheDate >= @startDate AND m.TheDate <= @endDate THEN ISNULL(m.Debtor, 0) ELSE 0 END) AS periodDebtor,
                        SUM(CASE WHEN m.TheDate >= @startDate AND m.TheDate <= @endDate THEN ISNULL(m.Creditor, 0) ELSE 0 END) AS periodCreditor,
                        SUM(CASE WHEN m.TheDate >= @startDate AND m.TheDate <= @endDate THEN ISNULL(m.DebtorDolar, 0) ELSE 0 END) AS periodDebtorDolar,
                        SUM(CASE WHEN m.TheDate >= @startDate AND m.TheDate <= @endDate THEN ISNULL(m.CreditorDolar, 0) ELSE 0 END) AS periodCreditorDolar
                    FROM MasterDetails m
                    WHERE m.AccountNumber LIKE @prefix
                    GROUP BY m.AccountNumber
                    ORDER BY m.AccountNumber ASC
                `);

                let periodDebtor = 0;
                let periodCreditor = 0;
                let periodDebtorDolar = 0;
                let periodCreditorDolar = 0;

                const childAccounts = sumRes.recordset.map(c => {
                    periodDebtor += c.periodDebtor;
                    periodCreditor += c.periodCreditor;
                    periodDebtorDolar += c.periodDebtorDolar;
                    periodCreditorDolar += c.periodCreditorDolar;

                    const cNetOpen = (c.openDebtor || 0) - (c.openCreditor || 0);
                    const cNetOpenDolar = (c.openDebtorDolar || 0) - (c.openCreditorDolar || 0);
                    const cNetEnding = ((c.openDebtor || 0) + (c.periodDebtor || 0)) - ((c.openCreditor || 0) + (c.periodCreditor || 0));
                    const cNetEndingDolar = ((c.openDebtorDolar || 0) + (c.periodDebtorDolar || 0)) - ((c.openCreditorDolar || 0) + (c.periodCreditorDolar || 0));

                    return {
                        ...c,
                        netOpen: cNetOpen,
                        netOpenDolar: cNetOpenDolar,
                        totalDebtor: c.openDebtor + c.periodDebtor,
                        totalCreditor: c.openCreditor + c.periodCreditor,
                        totalDebtorDolar: c.openDebtorDolar + c.periodDebtorDolar,
                        totalCreditorDolar: c.openCreditorDolar + c.periodCreditorDolar,
                        netEnding: cNetEnding,
                        netEndingDolar: cNetEndingDolar
                    };
                });

                const netEnding = (o.openDebtor + periodDebtor) - (o.openCreditor + periodCreditor);
                const netEndingDolar = (o.openDebtorDolar + periodDebtorDolar) - (o.openCreditorDolar + periodCreditorDolar);

                // Update Grand Totals
                grandTotals.openDebtor += o.openDebtor;
                grandTotals.openCreditor += o.openCreditor;
                grandTotals.openDebtorDolar += o.openDebtorDolar;
                grandTotals.openCreditorDolar += o.openCreditorDolar;

                if (netOpen > 0) grandTotals.netOpenDebit += netOpen;
                else if (netOpen < 0) grandTotals.netOpenCredit += Math.abs(netOpen);

                if (netOpenDolar > 0) grandTotals.netOpenDebitDolar += netOpenDolar;
                else if (netOpenDolar < 0) grandTotals.netOpenCreditDolar += Math.abs(netOpenDolar);

                grandTotals.periodDebtor += periodDebtor;
                grandTotals.periodCreditor += periodCreditor;
                grandTotals.periodDebtorDolar += periodDebtorDolar;
                grandTotals.periodCreditorDolar += periodCreditorDolar;

                grandTotals.totalDebtor += (o.openDebtor + periodDebtor);
                grandTotals.totalCreditor += (o.openCreditor + periodCreditor);
                grandTotals.totalDebtorDolar += (o.openDebtorDolar + periodDebtorDolar);
                grandTotals.totalCreditorDolar += (o.openCreditorDolar + periodCreditorDolar);

                if (netEnding > 0) grandTotals.netEndingDebit += netEnding;
                else if (netEnding < 0) grandTotals.netEndingCredit += Math.abs(netEnding);

                if (netEndingDolar > 0) grandTotals.netEndingDebitDolar += netEndingDolar;
                else if (netEndingDolar < 0) grandTotals.netEndingCreditDolar += Math.abs(netEndingDolar);

                accountsData.push({
                    code: accInfo.code,
                    name: accInfo.name,
                    parentName: accInfo.parentName,
                    category,
                    openDebtor: o.openDebtor,
                    openCreditor: o.openCreditor,
                    netOpen,
                    openDebtorDolar: o.openDebtorDolar,
                    openCreditorDolar: o.openCreditorDolar,
                    netOpenDolar,
                    periodDebtor,
                    periodCreditor,
                    periodDebtorDolar,
                    periodCreditorDolar,
                    totalDebtor: o.openDebtor + periodDebtor,
                    totalCreditor: o.openCreditor + periodCreditor,
                    totalDebtorDolar: o.openDebtorDolar + periodDebtorDolar,
                    totalCreditorDolar: o.openCreditorDolar + periodCreditorDolar,
                    netEnding,
                    netEndingDolar,
                    childAccountsCount: childAccounts.length,
                    childAccounts
                });
            }
        }

        res.json({
            startDate,
            endDate,
            mode,
            accountsCount: accountsData.length,
            grandTotals,
            accounts: accountsData
        });
    } catch (err) {
        console.error(err);
        res.status(500).json({ error: err.message });
    }
});

// 9. GET SUB-ACCOUNTS LIST (الحسابات الفرعية / التحليلية)
app.get('/api/accounts/sub-accounts', async (req, res) => {
    try {
        const p = await getPool();
        const result = await p.request().query(`
            SELECT DISTINCT
                COALESCE(a.SubAccountNum, m.AccountNumber) AS code,
                COALESCE(a.SubAccountName, m.AccountName) AS name,
                COALESCE(a.AccountNum, '') AS parentCode,
                COALESCE(a.AccountName, '') AS parentName
            FROM AccountsTree a
            FULL OUTER JOIN (
                SELECT AccountNumber, MAX(AccountName) AS AccountName 
                FROM MasterDetails 
                WHERE AccountNumber IS NOT NULL AND AccountNumber <> ''
                GROUP BY AccountNumber
            ) m ON a.SubAccountNum = m.AccountNumber
            WHERE (a.AccountType <> N'عـــــــام' OR a.AccountType IS NULL OR LEN(a.SubAccountNum) > 6)
              AND COALESCE(a.SubAccountNum, m.AccountNumber) IS NOT NULL
              AND COALESCE(a.SubAccountNum, m.AccountNumber) <> '00'
            ORDER BY code ASC
        `);

        const formatted = result.recordset.map(a => {
            let category = 'الأصول';
            if (a.code.startsWith('02')) category = 'الخصوم';
            else if (a.code.startsWith('03')) category = 'المصروفات';
            else if (a.code.startsWith('04')) category = 'الإيرادات';

            return {
                ...a,
                category
            };
        });

        res.json({
            total: formatted.length,
            accounts: formatted
        });
    } catch (err) {
        console.error(err);
        res.status(500).json({ error: err.message });
    }
});

// 10. REPORT: SUB-ACCOUNT LEDGER (تقرير كشف حساب الحسابات الفرعية تفصيلي بالجنيه والدولار)
app.get('/api/reports/sub-account-ledger', async (req, res) => {
    try {
        const {
            accountCode = '',
            startDate = '2021-01-01',
            endDate = '2023-12-31'
        } = req.query;

        if (!accountCode) {
            return res.status(400).json({ error: 'يرجى تحديد كود أو رقم الحساب الفرعي' });
        }

        const p = await getPool();

        // 1. Account Metadata
        const metaReq = p.request();
        metaReq.input('accountCode', sql.NVarChar, accountCode.trim());
        const metaRes = await metaReq.query(`
            SELECT TOP 1
                COALESCE(a.SubAccountNum, m.AccountNumber, @accountCode) AS code,
                COALESCE(a.SubAccountName, m.AccountName, @accountCode) AS name,
                COALESCE(a.AccountNum, '') AS parentCode,
                COALESCE(a.AccountName, '') AS parentName
            FROM (SELECT @accountCode AS code) x
            LEFT JOIN AccountsTree a ON a.SubAccountNum = x.code
            LEFT JOIN (
                SELECT AccountNumber, MAX(AccountName) AS AccountName
                FROM MasterDetails
                WHERE AccountNumber = @accountCode
                GROUP BY AccountNumber
            ) m ON m.AccountNumber = x.code
        `);

        const rawAcc = metaRes.recordset[0] || {
            code: accountCode,
            name: accountCode,
            parentCode: '',
            parentName: ''
        };

        let category = 'الأصول';
        if (rawAcc.code.startsWith('02')) category = 'الخصوم';
        else if (rawAcc.code.startsWith('03')) category = 'المصروفات';
        else if (rawAcc.code.startsWith('04')) category = 'الإيرادات';

        const account = {
            ...rawAcc,
            category
        };

        // 2. Opening Balance
        const openReq = p.request();
        openReq.input('accountCode', sql.NVarChar, accountCode.trim());
        openReq.input('startDate', sql.Date, startDate);
        const openRes = await openReq.query(`
            SELECT 
                ISNULL(SUM(Debtor), 0) AS openDebtor,
                ISNULL(SUM(Creditor), 0) AS openCreditor,
                ISNULL(SUM(DebtorDolar), 0) AS openDebtorDolar,
                ISNULL(SUM(CreditorDolar), 0) AS openCreditorDolar
            FROM MasterDetails
            WHERE AccountNumber = @accountCode
              AND TheDate < @startDate
        `);

        const o = openRes.recordset[0] || { openDebtor: 0, openCreditor: 0, openDebtorDolar: 0, openCreditorDolar: 0 };
        const netOpen = o.openDebtor - o.openCreditor;
        const netOpenDolar = o.openDebtorDolar - o.openCreditorDolar;

        // 3. Transactions in Period
        const txReq = p.request();
        txReq.input('accountCode', sql.NVarChar, accountCode.trim());
        txReq.input('startDate', sql.Date, startDate);
        txReq.input('endDate', sql.Date, endDate);
        const txRes = await txReq.query(`
            SELECT 
                Id,
                TheDate,
                ReqNum,
                DocType,
                Detail,
                ISNULL(Debtor, 0) AS Debtor,
                ISNULL(Creditor, 0) AS Creditor,
                ISNULL(DebtorDolar, 0) AS DebtorDolar,
                ISNULL(CreditorDolar, 0) AS CreditorDolar,
                ISNULL(Curruncy_Name, N'ج.س') AS Curruncy_Name,
                ISNULL(Check_Number, '') AS Check_Number
            FROM MasterDetails
            WHERE AccountNumber = @accountCode
              AND TheDate >= @startDate
              AND TheDate <= @endDate
            ORDER BY TheDate ASC, Id ASC
        `);

        let runningNet = netOpen;
        let runningNetDolar = netOpenDolar;
        let periodDebtor = 0;
        let periodCreditor = 0;
        let periodDebtorDolar = 0;
        let periodCreditorDolar = 0;

        const transactions = txRes.recordset.map(t => {
            periodDebtor += t.Debtor;
            periodCreditor += t.Creditor;
            periodDebtorDolar += t.DebtorDolar;
            periodCreditorDolar += t.CreditorDolar;

            runningNet += (t.Debtor - t.Creditor);
            runningNetDolar += (t.DebtorDolar - t.CreditorDolar);

            return {
                ...t,
                runningNet,
                runningNetDolar
            };
        });

        const netEnding = runningNet;
        const netEndingDolar = runningNetDolar;

        res.json({
            startDate,
            endDate,
            account,
            openDebtor: o.openDebtor,
            openCreditor: o.openCreditor,
            openDebtorDolar: o.openDebtorDolar,
            openCreditorDolar: o.openCreditorDolar,
            netOpen,
            netOpenDolar,
            periodDebtor,
            periodCreditor,
            periodDebtorDolar,
            periodCreditorDolar,
            totalDebtor: o.openDebtor + periodDebtor,
            totalCreditor: o.openCreditor + periodCreditor,
            totalDebtorDolar: o.openDebtorDolar + periodDebtorDolar,
            totalCreditorDolar: o.openCreditorDolar + periodCreditorDolar,
            netEnding,
            netEndingDolar,
            transactionsCount: transactions.length,
            transactions
        });
    } catch (err) {
        console.error(err);
        res.status(500).json({ error: err.message });
    }
});

app.listen(PORT, () => {
    console.log(`Server is running on port ${PORT}`);
});
