const ExcelJS = require('exceljs')

excelTest()

const data = [
  {
    uuid: '393fj839fr-fij39-3u9r398',
    name: 'Dispositivo Grillo 1',
    habitat: 'Grillero',
    category: 'Ambient Insectos',
    temperature: 23,
    humidity: 56,
  },
  {
    uuid: 'dfd903-f390kfo-eriu3',
    name: 'Dispositivo Grillo 2',
    habitat: 'Grillero',
    category: 'Ambient Insectos',
    temperature: 20,
    humidity: 40,
  },
  {
    uuid: 'mc3903-309r3-30493',
    name: 'Para sapos',
    habitat: 'Ambiente Sapo',
    category: 'Ambient Anfibios',
    temperature: 24,
    humidity: 71,
  },
]

async function excelTest() {
  const wb = new ExcelJS.Workbook()
  await wb.xlsx.readFile('log.xlsx')

  const ws = wb.getWorksheet(1)
  console.log(JSON.stringify(ws.getTables(), undefined, 2))
  const table = ws.getTable('Datos')

  data.forEach((device) => {
    console.log('agregando fila');
    table.addRow([
      device.uuid,
      device.name,
      device.habitat,
      device.category,
      device.temperature,
      device.humidity,
    ])
  })

  await wb.xlsx.writeFile('log.xlsx')
}
