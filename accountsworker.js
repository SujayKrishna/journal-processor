importScripts('jszip.js');
importScripts('xlsx.js');

const DR_ACCOUNT = "DR ACCOUNT";
const CR_ACCOUNT = "CR ACCOUNT";
const AMOUNT = "AMOUNT";

const LEDGER_HEADERS = [
	"DATE",
	"JV NO.",
	"DR/CR",
	"ACCOUNT",
	"PARTICULARS",
	"CHQ NO.",
	"DEBIT",
	"CREDIT",
	"BALANCE",
];

const TRIAL_BALANCE_HEADERS = [
	"ACCOUNT",
	"DEBIT",
	"CREDIT",
];

postMessage({type: "ready"});

let global_wb;
let allRows;
let allAccounts;
let trialBalances = {};
let buf;

function sheet_to_array_buffer(sheet, sheetName) {
	const wb = XLSX.utils.book_new();
	XLSX.utils.book_append_sheet(wb, sheet, sheetName);
	let binaryString = XLSX.write(wb, {bookType: 'xlsx', type: 'binary'});
	let buf = new ArrayBuffer(binaryString.length);
	let view = new Uint8Array(buf);
	for (let i=0; i < binaryString.length; i++) {
		view[i] = binaryString.charCodeAt(i) & 0xFF;
	}

	return buf;
}

onmessage = function (evt) {
	switch(evt.data.type) {
		case 'set_global_wb':
			global_wb = evt.data.workbook;
			allRows = global_wb.Sheet1;
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
			console.log(accountName);
			const filteredRows = allRows.filter(
				row => row[DR_ACCOUNT] === accountName || row[CR_ACCOUNT] === accountName
			);
			let balance = 0;
			const ledgerEntries = filteredRows.map(
				row => {
					const rowSubset = LEDGER_HEADERS.reduce((obj, header) => {
						obj[header] = row[header];
						return obj;
					}, {});
					const isDebit = row[DR_ACCOUNT] === accountName;
					const debitCreditStr = isDebit ? 'DR' : 'CR';
					const amount = row[AMOUNT];
					const debit = isDebit ? amount : 0;
					const credit = isDebit ? 0 : amount;

					balance += (isDebit ? -1 : 1) * amount;

					return {
						...rowSubset,
						ACCOUNT: accountName,
						"DR/CR": debitCreditStr,
						DEBIT: debit || '',
						CREDIT: credit || '',
						BALANCE: balance,
					};
				}
			);
			trialBalances[accountName] = balance;
			console.log('num filtered rows', filteredRows.length);
			const ledgerSheet = XLSX.utils.json_to_sheet(ledgerEntries, {
				header: LEDGER_HEADERS,
			});
			buf = sheet_to_array_buffer(ledgerSheet, `${accountName} Ledger`);
			postMessage({type: "save_excel", buffer: buf, fileName: `${accountName} Ledger`}, [buf]);
			break;
		case 'get_trial_balance':
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

			const trialBalanceEntries = [...allAccounts].map(account => {
				const balance = trialBalances[account];
				const debit = balance < 0 ? -balance : 0;
				const credit = balance > 0 ? balance : 0;

				return {
					ACCOUNT: account,
					DEBIT: debit || '',
					CREDIT: credit || '',
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
			trialBalanceSheet[`B${numRows + 2}`] = {f: `SUM(B2:B${trialBalanceEntries.length - 1})`};
			trialBalanceSheet[`C${numRows + 2}`] = {f: `SUM(C2:C${trialBalanceEntries.length - 1})`};
			buf = sheet_to_array_buffer(trialBalanceSheet, 'Trial Balance');
			postMessage({type: "save_excel", buffer: buf, fileName: 'Trial Balance'}, [buf]);
			break;
	}

};