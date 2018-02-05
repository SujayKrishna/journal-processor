let global_wb;
let global_accounts_list;
let excelParsingWorker;
const accountsWorker = new Worker('./accountsworker.js');
let accountsWorkerReady = false;
let pendingExcelData;
accountsWorker.onmessage = function(e) {
	switch(e.data.type) {
		case 'ready':
			accountsWorkerReady = true;
			if (pendingExcelData) {
				accountsWorker.postMessage({type: 'read_excel', data: pendingExcelData});
				pendingExcelData = null;
			}
			break;
		case 'ready_to_process':
			accountsWorker.postMessage({type: 'get_all_accounts'});
			show_actions();
			break;
		case 'all_accounts':
			global_accounts_list = e.data.accounts.sort();
			add_accounts();
			break;
		case 'save_excel':
			saveAs(new Blob([e.data.buffer], {type: "application/octet-stream"}), `${e.data.fileName}.xlsx`);
			break;
		case 'save_zip':
			saveAs(new Blob([e.data.buffer], {type: "application/octet-stream"}), `${e.data.fileName}.zip`);
			break;
	}
};

const reader = new FileReader();
reader.onload = function(e) {
	const data = e.target.result;
	if (accountsWorkerReady) {
		accountsWorker.postMessage({type: 'read_excel', data: data});
	} else {
		pendingExcelData = data;
	}
};

function add_accounts() {
	const accountsListSelect = document.getElementById('accounts_list');
	const optionItems = global_accounts_list.map(account => `<option value="${account}">${account}</option>`
	);
	accountsListSelect.innerHTML += optionItems.join('');
}

function show_actions() {
	const actionsContainer = document.getElementById('actions_container');
	actionsContainer.classList.remove('hide');
}

function process_journal(files) {
	let file = files.length && files[0];
	reader.readAsBinaryString(file);
}

(function() {
	var drop = document.getElementById('drop');
	if(!drop.addEventListener) return;

	function handleDrop(e) {
		e.stopPropagation();
		e.preventDefault();
		process_journal(e.dataTransfer.files);
	}

	function handleDragover(e) {
		e.stopPropagation();
		e.preventDefault();
		e.dataTransfer.dropEffect = 'copy';
	}

	drop.addEventListener('dragenter', handleDragover, false);
	drop.addEventListener('dragover', handleDragover, false);
	drop.addEventListener('drop', handleDrop, false);
})();

(function() {
	let xlf = document.getElementById('xlf');
	if(!xlf.addEventListener) return;
	function handleFile(e) { process_journal(e.target.files); }
	xlf.addEventListener('change', handleFile, false);
})();

(function() {
	const accountsListSelect = document.getElementById('accounts_list');
	accountsListSelect.addEventListener('change', function changeHandler() {
		const accountName = accountsListSelect.value;
		if (!accountName) {
			return false;
		}
		accountsWorker && accountsWorker.postMessage({type: 'get_account_ledger', accountName});
	});
})();

(function() {
	const trialBalanceLink = document.getElementById('trial_balance');
	trialBalanceLink.addEventListener('click', function clickHandler(e) {
		e.stopPropagation();
		e.preventDefault();
		accountsWorker && accountsWorker.postMessage({type: 'get_trial_balance'});
	});
})();

(function() {
	const zipAllLedgersLink = document.getElementById('zip_all_ledgers');
	zipAllLedgersLink.addEventListener('click', function clickHandler(e) {
		e.stopPropagation();
		e.preventDefault();
		accountsWorker && accountsWorker.postMessage({type: 'get_zip_of_all_ledgers'});
	});
})();

(function() {
	const zipAllLink = document.getElementById('zip_all');
	zipAllLink.addEventListener('click', function clickHandler(e) {
		e.stopPropagation();
		e.preventDefault();
		accountsWorker && accountsWorker.postMessage({type: 'get_zip_of_all'});
	});
})();
