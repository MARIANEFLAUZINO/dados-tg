const Excel = require('exceljs');
const path = require('path');

const data = {
    years: [2019, 2020, 2021],
    months: ['Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho', 'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro'],
    sellers: ['Lucimar Sasso', 'Mariane Flauzino', 'Mateus Alves'],
    regions: ['Sul', 'Suldeste', 'Centro-Oeste', 'Norte', 'Nordeste'],
    products: [{ item: 'Pacote Office', price: 450, acc: 35 }, { item: 'Power BI', price: 500, acc: 40 }, { item: 'Microsoft Azure', price: 600, acc: 50 }],
    paymentMethods: ['Cartão de Débito', 'Cartão de Crédito', 'Boleto']
}


function randomArrayIndex(size) {
    return Math.floor(Math.random() * size)
}

function getRandomData(key) {
    const index = randomArrayIndex(data[key].length)
    const random = data[key][index]
    return random
}

function getRandomDate(month, year) {
    const lastDayOfMonth = new Date(year, month, 0).getDate()
    const randomDay = Math.floor(Math.random() * lastDayOfMonth) + 1
    return `${randomDay < 10 ? '0' : ''}${randomDay}/${month < 10 ? '0' : ''}${month}/${year}`
}

function calcProductPrice(product, year) {
    let { price, item } = product
    switch (year) {
        case 2020:
            price += product.acc
            break
        case 2021:
            price += product.acc * 2
            break
        default:
            price = product.price
            break
    }
    return { price, item }
}

const fakeData = []
const limit = 80 * 1000
let index = 0

while (index < limit) {
    const row = {
        year: getRandomData('years'),
        month: getRandomData('months'),
        seller: getRandomData('sellers'),
        region: getRandomData('regions'),
        paymentMethod: getRandomData('paymentMethods')
    }
    const product = calcProductPrice(getRandomData('products'), row.year)
    row.product = product.item
    row.price = product.price
    row.date = getRandomDate(data.months.findIndex(m => m == row.month) + 1, row.year)
    fakeData.push(row)
    index++
}

const workbook = new Excel.Workbook()
const workSheet = workbook.addWorksheet('Dados')

// Definir cabeçalhos
workSheet.columns = [
    { header: 'Vendedor', key: 'seller', width: 20 },
    { header: 'Produto', key: 'product', width: 10 },
    { header: 'Preço', key: 'price', width: 10 },
    { header: 'Forma de Pagamento', key: 'paymentMethod', width: 20 },
    { header: 'Data', key: 'date', width: '20' },
    { header: 'Região', key: 'region', width: 20 },
    { header: 'Mês', key: 'month', width: 20 },
    { header: 'Ano', key: 'year', with: 20 }
];

// Preencher a planilha com os dados
fakeData.forEach((row) => {
    workSheet.addRow(row);
});

// Salvar o arquivo Excel
const filePath = path.join(__dirname, 'dados.xlsx');

workbook.xlsx.writeFile(filePath)
    .then(() => {
        console.log(`Arquivo Excel exportado com sucesso para: ${filePath}`);
    })
    .catch((error) => {
        console.error('Erro ao exportar para Excel:', error);
    });