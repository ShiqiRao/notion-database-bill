import { Client } from "@notionhq/client";
import xlsx from 'xlsx';
import { auth, databaseId } from "./config.js";


const { readFile, utils } = xlsx;
var workbook = readFile('./wx.csv', { cellDates: true });
var first_sheet_name = workbook.SheetNames[0];
var worksheet = workbook.Sheets[first_sheet_name];
var cursor = 18;
const notion = new Client({ auth })
while (readLine(cursor, worksheet)) {
    cursor++;
}



async function addItem(date, title, income, expenses) {
    try {
        const response = await notion.pages.create({
            parent: {
                database_id: databaseId,
            },
            properties: {
                '项目名': {
                    title: [
                        {
                            text: {
                                content: title,
                            },
                        },
                    ],
                },
                '日期': {
                    date: {
                        start: date,
                    },
                },
                '收入': {
                    number: income,
                },
                '支出': {
                    number: expenses,
                },
            },
        });
        console.log(response);
    } catch (error) {
        console.error(error);
    }
}

function moneyToNumber(money) {
    var yuanStart = money.indexOf("¥");
    if (yuanStart == 0) {
        money = money.substr(1, money.length);
    }
    return parseFloat(money.replace(',\g', ''));
}

function readLine(lineNumber, worksheet) {
    var firstCell = worksheet['A' + lineNumber];
    if (firstCell != undefined) {
        var date = new Date(worksheet['A' + lineNumber].v);
        var date_string = date.toISOString();
        var title = worksheet['C' + lineNumber].v;
        var type = worksheet['E' + lineNumber].v;
        var amount = worksheet['F' + lineNumber].v;
        var income = 0.0;
        var expenses = 0.0;
        if (type == '支出') {
            income = moneyToNumber(amount);
        }
        if (type == '收入') {
            expenses = moneyToNumber(amount);
        }
        console.log(`${date_string} ${title} ${type} ${income} ${expenses}`);
        addItem(date_string, title, income, expenses);
        return true;
    }
    return false;
}

async function getPagesFromNotionDatabase() {
    const pages = []
    let cursor = undefined
    while (true) {
        const { results, next_cursor } = await notion.databases.query({
            database_id: databaseId,
            start_cursor: cursor,
        })
        pages.push(...results)
        if (!next_cursor) {
            break
        }
        cursor = next_cursor
    }
    console.log(pages);
    console.log(`${pages.length} pages successfully fetched.`)
    return pages.map(page => {
        return {
            pageId: page.id,
        }
    })
}