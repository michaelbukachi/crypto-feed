const fs = require('fs');
const axios = require('axios');
const Excel = require('exceljs');
const moment = require('moment');

const FILENAME = 'out.xlsx';
const APIS = ['Bitstamp', 'Hitbtc', 'Gate', 'exmo', 'Yobit'];

// Add bitsmap coins in the following array
const BITSMAP_COINS = ['btcusd', 'eurusd', 'xrpusd', 'ltcusd', 'ethusd', 'bchusd'];

let workbook;

// Create or load excel file
function createExcel() {
    if (!fs.existsSync(FILENAME)) {
        workbook = new Excel.Workbook();
        workbook.creator = 'Me';
        workbook.created = new Date();
        APIS.forEach(function (api) {
            let worksheet = workbook.addWorksheet(api);
            worksheet.columns = [
                {header: 'Coin name', key: 'name', width: 10},
                {header: 'Ask', key: 'ask'},
                {header: 'Bid', key: 'bid'},
                {header: 'Last traded price', key: 'last', width: 20},
                {header: 'Timestamp', key: 'timestamp', width: 15},
            ];
        });
        workbook.xlsx.writeFile(FILENAME)
            .then(function () {
                console.log("File created");
                startFetching();
            });
    } else {
        console.log('File exists. Opening..');
        workbook = new Excel.Workbook();
        workbook.xlsx.readFile(FILENAME)
            .then(function () {
                console.log("File loaded");
                startFetching();
            });
    }
}

function addToExcel(sheet, coin, ask, bid, last, timestamp) {
    coin = coin.replace('_', '');
    let worksheet = workbook.getWorksheet(sheet);
    let modified = false;
    for (let i = 0; i < worksheet.rowCount; i++) {
        let row = worksheet.getRow(i);
        if (row.values[1] === coin) {
            row.values = [coin, ask, bid, last, moment(timestamp).format('h:mm:ss a')];
            row.commit();
            modified = true;
            break;
        }
    }
    if (!modified) {
        worksheet.addRow([coin, ask, bid, last, moment(timestamp).format('h:mm:ss a')]);
    }
}

function startBitstamp() {
    let _request = async function (coin) {
        try {
            const response = await axios.get('https://www.bitstamp.net/api/v2/ticker/' + coin);
            let data = response.data;
            console.log('Received data from bitstamp...');
            addToExcel(APIS[0], coin, data.ask, data.bid, data.last, moment.unix(data.timestamp));
        } catch (error) {
            console.error(error);
        }
    };

    let index = 0;
    setInterval(function () {
        _request(BITSMAP_COINS[index]);
        index++;

        if (index === BITSMAP_COINS.length) {
            index = 0;
        }
    }, 3000);
}

async function startHitbtc() {
    try {
        const response = await axios.get('https://api.hitbtc.com/api/2/public/symbol');
        let data = response.data;
        let coins = [];
        for (let i = 0; i < data.length; i++) {
            let id = data[i].id;
            if (id.endsWith('USD')) {
                coins.push(id);
            }
        }

        let _request = async function (coin) {
            try {
                const response = await axios.get('https://api.hitbtc.com/api/2/public/ticker/' + coin);
                let data = response.data;
                console.log('Received data from hitbtc...');
                addToExcel(APIS[1], coin.toLowerCase(), data.ask, data.bid, data.last, data.timestamp);
            } catch (error) {
                console.error(error);
            }
        };

        let index = 0;
        setInterval(function () {
            _request(coins[index]);
            index++;

            if (index === coins.length) {
                index = 0;
            }
        }, 3000);

    } catch (error) {
        console.error(error);
    }
}

function startGate() {
    let _request = async function () {
        try {
            const response = await axios.get('http://data.gate.io/api2/1/tickers');
            console.log('Received data from gate...');
            let data = response.data;
            for (let key in data) {
                if (key.endsWith('usdt')) {
                    let ticker = data[key];
                    addToExcel(APIS[2], key, ticker.lowestAsk, ticker.highestBid, ticker.last, new Date());
                }
            }
        } catch (error) {
            console.error(error);
        }
    };

    setInterval(function () {
        _request();
    }, 3000);
}

function startExmo() {
    let _request = async function () {
        try {
            const response = await axios.get('https://api.exmo.com/v1/ticker/');
            console.log('Received data from exmo...');
            let data = response.data;
            for (let key in data) {
                if (key.endsWith('USD')) {
                    let ticker = data[key];
                    addToExcel(APIS[3], key.toLowerCase(), ticker.sell_price, ticker.buy_price, ticker.last_trade, moment.unix(ticker.updated));
                }
            }
        } catch (error) {
            console.error(error);
        }
    };

    setInterval(function () {
        _request();
    }, 3000);
}

async function startYobit() {
    try {
        const response = await axios.get('https://yobit.net/api/3/info');
        let data = response.data;
        let coins = [];
        for (let key in data.pairs) {
            if (key.endsWith('usd')) {
                coins.push(key);
            }
        }

        let _request = async function (coin) {
            try {
                const response = await axios.get('https://yobit.net/api/3/ticker/' + coin);
                let data = response.data;
                data = data[coin];
                console.log('Received data from yobit...');

                addToExcel(APIS[4], coin, data.sell, data.buy, data.last, moment.unix(data.updated));
            } catch (error) {
                console.error(error);
            }
        };

        let index = 0;
        setInterval(function () {
            _request(coins[index]);
            index++;

            if (index === coins.length) {
                index = 0;
            }
        }, 3000);

    } catch (error) {
        console.error(error);
    }
}

function startFetching() {
    startBitstamp();
    startHitbtc();
    startGate();
    startExmo();
    startYobit();
}

function start() {
    console.log('Starting....');
    createExcel();
}

process.on('SIGINT', function () {
    workbook.xlsx.writeFile(FILENAME)
        .then(function () {
            console.log("Exiting..");
            process.exit();
        });
});

start();
