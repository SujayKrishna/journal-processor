importScripts('jszip.js');
importScripts('xlsx.js');

const EXCEL_PARSING_OPTIONS = {
	type: 'binary', 
	cellDates: true, 
	cellNF: true,
	cellStyles: true,
};

const EXCEL_WRITING_OPTIONS = {
	bookType: 'xlsx', 
	type: 'binary', 
	cellDates: true,
};

const SHEET_OPTIONS = {
	Sheet1: {range: 2, raw: true},
};

const DR_ACCOUNT = "DR ACCOUNT";
const CR_ACCOUNT = "CR ACCOUNT";
const AMOUNT = "AMOUNT";

const ACCOUNT = "ACCOUNT";
const ACCOUNT_FULL_NAME = "ACCOUNT FULL NAME";
const LEDGER_ACCOUNT = "LEDGER ACCOUNT";

const LEDGER_HEADERS = [
	"DATE",
	"JV NO.",
	"DR/CR",
	"ACCOUNT",
	"ACCOUNT FULL NAME",
	"PARTICULARS",
	"CHQ NO.",
	"DEBIT",
	"CREDIT",
	"BALANCE",
];

const LEDGER_NUMBER_COLUMNS = [
	'H', // DEBIT
	'I', // CREDIT
	'J', // BALANCE
];

const TRIAL_BALANCE_HEADERS = [
	"ACCOUNT",
	"ACCOUNT FULL NAME",
	"DEBIT",
	"CREDIT",
];

const TRIAL_BALANCE_NUMBER_COLUMNS = [
	'C', // DEBIT
	'D', // CREDIT
];

const LEDGER_SHEET_OPTIONS = {
	header: LEDGER_HEADERS,
	cellDates: true,
};

let ledgerCols;
let trialBalanceCols;

let global_wb;
let allRows;
let accountShortToFull;
let allAccounts;
let trialBalances = {};
let buf;
let zip;

function workbook_to_json(workbook) {
	let result = {};
	workbook.SheetNames.forEach(function(sheetName) {
		let sheetAsJSON = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], SHEET_OPTIONS[sheetName]);
		if (sheetAsJSON.length){
			result[sheetName] = sheetAsJSON;	
		} 
	});
	return result;
}

function sheet_to_array_buffer(sheet, sheetName) {
	const wb = XLSX.utils.book_new();
	XLSX.utils.book_append_sheet(wb, sheet, sheetName);
	let binaryString = XLSX.write(wb, EXCEL_WRITING_OPTIONS);
	let buf = new ArrayBuffer(binaryString.length);
	let view = new Uint8Array(buf);
	for (let i=0; i < binaryString.length; i++) {
		view[i] = binaryString.charCodeAt(i) & 0xFF;
	}

	return buf;
}

function set_number_format_for_cells(sheet, numRows, columns) {
	for (let i = 2; i <= numRows; i++) {
		for (let column of columns) {
			if (!sheet[`${column}${i}`]) {
				continue;
			}
			XLSX.utils.cell_set_number_format(sheet[`${column}${i}`], "0.00");
		}
	}
}

function getTrialBalanceSheet() {
	let accountsToProcess = new Set();
	allAccounts.forEach(account => {
		if (!(account in trialBalances)) {
			trialBalances[account] = 0;
			accountsToProcess.add(account);
		}
	});

	allRows.forEach(row => {
		const debitAccount = row[DR_ACCOUNT];
		const creditAccount = row[CR_ACCOUNT];
		const amount = +row[AMOUNT];

		if (accountsToProcess.has(debitAccount)) {
			trialBalances[debitAccount] -= amount;
		}

		if (accountsToProcess.has(creditAccount)) {
			trialBalances[creditAccount] += amount;
		}
	});

	const trialBalanceEntries = [...allAccounts].sort().map(account => {
		const balance = trialBalances[account];
		const debit = balance < 0 ? -balance : 0;
		const credit = balance > 0 ? balance : 0;

		return {
			[TRIAL_BALANCE_HEADERS[0]]: account,
			[TRIAL_BALANCE_HEADERS[1]]: accountShortToFull[account] || account,
			[TRIAL_BALANCE_HEADERS[2]]: debit || '',
			[TRIAL_BALANCE_HEADERS[3]]: credit || '',
		};
	});
	const numRows = trialBalanceEntries.length + 1;
	trialBalanceEntries.push({});
	trialBalanceEntries.push({
		ACCOUNT: 'TOTAL',
	});
	const trialBalanceSheet = XLSX.utils.json_to_sheet(trialBalanceEntries, {
		header: TRIAL_BALANCE_HEADERS,
	});
	trialBalanceSheet[`C${numRows + 2}`] = {f: `SUM(C2:C${trialBalanceEntries.length - 1})`};
	trialBalanceSheet[`D${numRows + 2}`] = {f: `SUM(D2:D${trialBalanceEntries.length - 1})`};
	trialBalanceSheet['!cols'] = JSON.parse(JSON.stringify(trialBalanceCols));
	set_number_format_for_cells(trialBalanceSheet, numRows + 2, TRIAL_BALANCE_NUMBER_COLUMNS);
	return sheet_to_array_buffer(trialBalanceSheet, 'Trial Balance');
}

function getLedgersForAccounts(accounts) {
	const filteredRows = allRows.filter(row => 
		accounts.includes(row[DR_ACCOUNT]) || accounts.includes(row[CR_ACCOUNT])
	);
	const accountNameToLedgerEntries = accounts.reduce((map, account) => {
		map[account] = [];
		return map;
	}, {});

	const accountNameToBalance = accounts.reduce((map, account) => {
		map[account] = 0;
		return map;
	}, {});

	filteredRows.forEach(row => {
		const rowSubset = LEDGER_HEADERS.reduce((obj, header) => {
			obj[header] = row[header];
			return obj;
		}, {});
		const debitAccount = row[DR_ACCOUNT];
		const creditAccount = row[CR_ACCOUNT];
		const amount = row[AMOUNT];
		[debitAccount, creditAccount].filter(account => accounts.includes(account)).forEach(account => {
			const isDebit = debitAccount === account;
			const debitCreditStr = isDebit ? 'DR' : 'CR';
			const debit = isDebit ? amount : 0;
			const credit = isDebit ? 0 : amount;

			accountNameToBalance[account] += (isDebit ? -1 : 1) * amount;

			accountNameToLedgerEntries[account].push({
				...rowSubset,
				ACCOUNT: account,
				[ACCOUNT_FULL_NAME]: accountShortToFull[account] || account,
				"DR/CR": debitCreditStr,
				DEBIT: debit || '',
				CREDIT: credit || '',
				BALANCE: accountNameToBalance[account],
			});
		});
	});

	trialBalances = Object.assign(trialBalances, accountNameToBalance);
	return accountNameToLedgerEntries;
}

function getLedgerSheet(ledgerEntries, sheetName) {
	const ledgerSheet = XLSX.utils.json_to_sheet(ledgerEntries, LEDGER_SHEET_OPTIONS);
	ledgerSheet['!cols'] = JSON.parse(JSON.stringify(ledgerCols));
	set_number_format_for_cells(ledgerSheet, ledgerEntries.length + 1, LEDGER_NUMBER_COLUMNS);
	return sheet_to_array_buffer(ledgerSheet, sheetName);
}

function getAllLedgersZip() {
	const accounts = [...allAccounts];
	const allLedgerEntries = getLedgersForAccounts(accounts);

	const zip = new JSZip();

	accounts.forEach(account => {
		const sheetName = `${account} Ledger`;
		const ledgerEntries = allLedgerEntries[account];
		zip.file(`${sheetName}.xlsx`, getLedgerSheet(ledgerEntries, sheetName));
	});

	return zip;
}

function getLedgerSheetForSingleAccount(account) {
	const ledgerEntries = getLedgersForAccounts([account])[account];
	const sheetName = `${account} Ledger`;
	return getLedgerSheet(ledgerEntries, sheetName);
}

onmessage = function (evt) {
	switch(evt.data.type) {
		case 'read_excel':
			const excelWorkbook = XLSX.read(evt.data.data, EXCEL_PARSING_OPTIONS);
			global_wb = workbook_to_json(excelWorkbook);
			allRows = global_wb.Sheet1;
			allRows.sort((row1, row2) => 
				row1['DATE'] - row2['DATE']
			);
			accountShortToFull = global_wb.Sheet2.reduce((map, row) => {
				map[row[ACCOUNT]] = row[LEDGER_ACCOUNT];
				return map;
			});
			const journalCols = excelWorkbook.Sheets.Sheet1['!cols'];
			const accountCols = excelWorkbook.Sheets.Sheet2['!cols'];
			ledgerCols = [
				journalCols[0],
				journalCols[1],
				journalCols[1],
				accountCols[0],
				accountCols[1],
				journalCols[4],
				journalCols[5],
				journalCols[6],
				journalCols[6],
				journalCols[6],
			];
			trialBalanceCols = [
				accountCols[0],
				accountCols[1],
				journalCols[6],
				journalCols[6],
			];
			postMessage({type: 'ready_to_process'});
			break;
		case 'get_all_accounts':
			if (typeof allAccounts == 'undefined') {
				allAccounts = new Set();
				allRows.forEach(row => {
					allAccounts.add(row[DR_ACCOUNT]);
					allAccounts.add(row[CR_ACCOUNT]);
				});
			}
			postMessage({type: "all_accounts", accounts: [...allAccounts]});
			break;
		case 'get_account_ledger':
			const accountName = evt.data.accountName;
			buf = getLedgerSheetForSingleAccount(accountName);
			postMessage({type: "save_excel", buffer: buf, fileName: `${accountName} Ledger`}, [buf]);
			break;
		case 'get_trial_balance':
			buf = getTrialBalanceSheet();
			postMessage({type: "save_excel", buffer: buf, fileName: 'Trial Balance'}, [buf]);
			break;
		case 'get_zip_of_all_ledgers':
			zip = getAllLedgersZip();
			buf = zip.generate({type:"arraybuffer"});
			postMessage({type: 'save_zip', buffer: buf, fileName: 'All Ledgers'}, [buf]);
			break;
		case 'get_zip_of_all':
			zip = getAllLedgersZip();
			const trialBalanceSheet = getTrialBalanceSheet();
			zip.file("TrialBalance.xlsx", trialBalanceSheet)
			buf = zip.generate({type:"arraybuffer"});
			postMessage({type: 'save_zip', buffer: buf, fileName: 'All Ledgers & Trial Balance'}, [buf]);
			break;
	}

};

postMessage({type: "ready"});