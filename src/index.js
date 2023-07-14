import { init } from 'excelize-wasm'
import wasm from './excelize.wasm.gz'

kintone.events.on('app.record.detail.show', async (event) => {
  const excelize = await init(wasm)
  URL.revokeObjectURL(wasm)
  const fileInfo = event.record.Template.value[0]
  const generatie = document.createElement('button')
  kintone.app.record.getHeaderMenuSpaceElement().appendChild(generatie)
  generatie.id = 'generatie'
  generatie.innerText = 'export'
  generatie.onclick = async () => {
    try {
      const fileUrl = `/k/v1/file.json?fileKey=${fileInfo.fileKey}`
      const response = await fetch(fileUrl, {
        method: 'GET',
        headers: {
          'X-Requested-With': 'XMLHttpRequest',
        },
      })
      const buff = await response.arrayBuffer()
      const newFile = excelize.OpenReader(new Uint8Array(buff))
      if (newFile.error) throw new Error(newFile.error)
      const start = 25
      const sheet = '御見積もり'
      const setCellValue = (cell, value) => {
        const { error } = newFile.SetCellValue(sheet, cell, value)
        if (error) throw new Error(error)
      }
      event.record.Table.value.forEach((row, index) => {
        const i = index + start
        const { error } = newFile.DuplicateRow(sheet, i)
        if (error) throw new Error(error)
        setCellValue(`B${i}`, row.value.content.value)
        setCellValue(`Q${i}`, row.value.quantity.value)
        setCellValue(`S${i}`, row.value.unit.value)
        setCellValue(`U${i}`, row.value.price.value)
        setCellValue(`X${i}`, row.value.total.value)
      })
      const { buffer, error } = newFile.WriteToBuffer()
      if (error) throw new Error(error)
      const link = document.createElement('a')
      link.download = 'Book1.xlsx'
      link.href = URL.createObjectURL(
        new Blob([buffer], {
          type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        }),
      )
      link.click()
      URL.revokeObjectURL(link.href)
    } catch (error) {
      console.error(error)
    }
  }
  return event
})
