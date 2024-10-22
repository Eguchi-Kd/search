const REDIRECT_URI = 'https://eguchi-kd.github.io/search/';
const RANGE = 'sheet1!A2:Z'; // 1行目は除外して2行目以降を対象とする

// 認証とAPIの読み込み
const DISCOVERY_DOC = 'https://sheets.googleapis.com/$discovery/rest?version=v4';
const SCOPES = 'https://www.googleapis.com/auth/spreadsheets.readonly';

let tokenClient;
let gapiInited = false;
let gisInited = false;
let checkboxStates = []; // 追加: チェックボックスの状態を保持する配列

// 認証ボタンを表示
document.getElementById('authorize_button').onclick = handleAuthClick;
document.getElementById('signout_button').onclick = handleSignoutClick;

// GAPIライブラリの読み込み
function gapiLoaded() {
    gapi.load('client', initializeGapiClient);
}

async function initializeGapiClient() {
    await gapi.client.init({
        discoveryDocs: [DISCOVERY_DOC],
    });
    gapiInited = true;
    maybeEnableButtons();
}

// Google Identity Servicesのライブラリ読み込み
function gisLoaded() {
    tokenClient = google.accounts.oauth2.initTokenClient({
        client_id: CLIENT_ID,
        scope: SCOPES,
        callback: '', // トークンのコールバック
        redirect_uri: REDIRECT_URI // 修正: REDIRECT_URIを使用
    });
    gisInited = true;
    maybeEnableButtons();
}

function maybeEnableButtons() {
    if (gapiInited && gisInited) {
        document.getElementById('authorize_button').style.display = 'block';
    }
}

// 認証処理
function handleAuthClick() {
    tokenClient.callback = async (resp) => {
        if (resp.error !== undefined) {
            throw (resp);
        }
        document.getElementById('authorize_button').style.display = 'none';
        document.getElementById('signout_button').style.display = 'block';
        await listMajors();
    };

    if (gapi.client.getToken() === null) {
        tokenClient.requestAccessToken({ prompt: 'consent' });
    } else {
        tokenClient.requestAccessToken({ prompt: '' });
    }
}

function handleSignoutClick() {
    const token = gapi.client.getToken();
    if (token !== null) {
        google.accounts.oauth2.revoke(token.access_token);
        gapi.client.setToken('');
        document.getElementById('authorize_button').style.display = 'block';
        document.getElementById('signout_button').style.display = 'none';
    }
}

// スプレッドシートのデータを取得
async function listMajors() {
    let response;
    try {
        response = await gapi.client.sheets.spreadsheets.values.get({
            spreadsheetId: SPREADSHEET_ID,
            range: RANGE,
        });
    } catch (err) {
        console.log(err.message);
        return;
    }

    const range = response.result;
    if (!range || !range.values || range.values.length == 0) {
        console.log('No values found.');
        return;
    }

    console.log('Excel data loaded successfully:', range.values);

    // チェックボックスの状態を初期化
    checkboxStates = new Array(range.values.length).fill(false);

    // ワード検索
    document.getElementById('wordSearchBtn').addEventListener('click', () => {
        const searchWord = document.getElementById("searchWord").value.toLowerCase();
        const tableBody = document.getElementById("tableBody");
        tableBody.innerHTML = ''; // 検索結果をクリア
        const results = []; // 検索結果を格納する配列
        range.values.forEach((row, rowIndex) => {
            row.forEach(cell => {
                if (typeof cell === 'string' && cell.toLowerCase().includes(searchWord)) {
                    results.push({ cellData: cell, rowIndex });
                }
            });
        });
        displayResults(results);
        updateSearchTime();
    });

    // 人数選択時に検索ワードに自動入力して検索
    document.getElementById('peopleCount').addEventListener('change', () => {
        const peopleCount = document.getElementById("peopleCount").value;
        const searchWord = peopleCount ? `${peopleCount}人` : ""; // 修正: テンプレートリテラルをバッククオートで囲む
        document.getElementById("searchWord").value = searchWord; // 検索ワードに自動入力

        // ワード検索を自動で実行
        document.getElementById("wordSearchBtn").click();
    });
}

// 検索結果を表示
function displayResults(results) {
    const tableBody = document.getElementById("tableBody");
    tableBody.innerHTML = ''; // 検索結果をクリア

    results.forEach(({ cellData, rowIndex }) => {
        const tr = document.createElement("tr");

        const tdData = document.createElement("td");
        tdData.textContent = cellData;
        tr.appendChild(tdData);

        const tdCheckbox = document.createElement("td");
        const checkbox = document.createElement("input");
        checkbox.type = "checkbox";
        checkbox.checked = checkboxStates[rowIndex] || false; // 状態を保持
        checkbox.addEventListener('change', () => {
            checkboxStates[rowIndex] = checkbox.checked; // 状態を更新
        });
        tdCheckbox.appendChild(checkbox);
        tr.appendChild(tdCheckbox);

        tableBody.appendChild(tr);
    });
}

// 現在の日時を更新
function updateSearchTime() {
    const now = new Date();
    document.getElementById("searchTime").textContent = now.toLocaleString();
}
