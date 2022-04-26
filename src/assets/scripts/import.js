import { WorkbookDef, SheetDef, ColumnDef } from './excel'
/**
 * @returns WorkbookDef
 * data: 数组对象，对象里面包含数据data和sheet表名dataName
 * data 是一个数组对象，如 [{ label: '编号', field: 'no', required: true }, {...}]，label 用于 excel 的表头字段名，field 用于 传给后台的字段名，required 表示是否一定为空。
 * dataName 是生成 excel 表名的一部分
 */
function getWbDef(data, rowHandler) {
  let sheet = []
  data.forEach(ele => {
    sheet.push(
      new SheetDef(
	      dataName,
	      data.map(item => {
	        return new ColumnDef(item.label, item.field, {
	          required: item.required
	        })
	      }),
	      rowHandler
      )
    )
  })
  return new WorkbookDef(sheet)
}

/**
 * 资产数据的导入和导入模板下载
 */
export default {
  /**
   * 下载导入的模板文件
   * @param {string} dataName 多级名称写法: 服务器/国产化服务器
   */
  downloadTemplate(data, dataName) {
    getWbDef(data).write([], dataName.replace('/', '_') + '-数据导入模板')
  },
  /**
   *
   * @param {string} data 数组对象
   * @param {File} file 选择的文件对象
   */
  async readData(data, file, rowHandler) {
    return getWbDef(data, rowHandler).read(file)
  },
  
  /**
   * 导出excel
   * @param {Array} data 写入excel的数据
   * { name: dataName, rows: data }：数据格式
   * 传入excel表的数据格式必须是 [ { name: dataName, rows: data }, { name: dataName, rows: data }, ... ]，数组里面有几个对象就有几个sheet表
   */
  export(data, dataName) {
    getDef(dataName).write(data, dataName)
  }
}
