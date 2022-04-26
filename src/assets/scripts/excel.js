import XLSX from 'xlsx'

function downloadData(data, filename) {
  const link = document.createElement('a')
  document.body.appendChild(link)
  link.href = URL.createObjectURL(data)
  link.download = filename
  link.click()

  setTimeout(() => {
    document.body.removeChild(link)
  }, 1000)
}

// 字符串转ArrayBuffer
function string2arraybuffer(s) {
  const buf = new ArrayBuffer(s.length)
  const view = new Uint8Array(buf)
  for (let i = 0; i !== s.length; ++i) view[i] = s.charCodeAt(i) & 0xff
  return buf
}

/**
 * 列的枚举类型定义
 */
class ColumnEnum {
  /**
   *
   * @param label
   * @param {any} value
   */
  constructor(label, value) {
    this.label = label
    this.value = value
  }
}

/**
 * 列定义
 */
class ColumnDef {
  /**
   *
   * @param {string} name 列的标题
   * @param {string} field 列的字段
   * @param {object} [options] 列选项
   * @param {bool} [options.required=true] 列是否必填的。当其为 true 时，其值不能为空
   * @param {string} [options.type='string'] 列的类型，可选值见 ColumnDef.types
   * @param {ColumnEnum[]} [options.enums] 列的枚举定义
   * @param {function({value: any, column: ColumnDef})} [options.writeParser] 列的写处理函数
   * @param {function({value: any, column: ColumnDef})} [options.readParser] 列的读处理函数
   */
  constructor(name, field, options) {
    this.name = name
    this.field = field

    this.options = Object.assign(
      {
        required: true,
        type: ColumnDef.types.STRING,
        enums: []
      },
      options
    )
  }

  /**
   * 处理导入时的数据
   * @param value
   * @returns {string|number|*}
   */
  parseReadValue(value) {
    if (this.options.readParser) {
      value = this.options.readParser({
        value,
        column: this
      })
    }
    const isEmpty = value === undefined || value === null || value === ''
    if (this.options.required && isEmpty) {
      throw new Error('值不能为空')
    }
    if (!isEmpty) {
      switch (this.options.type) {
        case ColumnDef.types.DATE:
        case ColumnDef.types.DATETIME:
          value = this._parseReadDate(value)
          break
        case ColumnDef.types.NUMBER:
          value = parseInt(value)
          break
        default:
          break
      }
      if (this.options.enums.length && [ColumnDef.types.DATE, ColumnDef.types.DATETIME].indexOf(this.options.type) === -1) {
        value = this._parseReadEnum()
      }
    }
    return value
  }

  /**
   * 处理导出时的数据
   * @param value
   * @returns {string|number|*}
   */
  parseWriteValue(value) {
    if (this.options.writeParser) {
      value = this.options.writeParser({
        value,
        column: this
      })
    }
    const isEmpty = value === undefined || value === null || value === ''
    if (!isEmpty) {
      switch (this.options.type) {
        case ColumnDef.types.DATE:
        case ColumnDef.types.DATETIME:
          value = this._parseWriteDate(value)
          break
        default:
          break
      }
      if (this.options.enums.length && [ColumnDef.types.DATE, ColumnDef.types.DATETIME].indexOf(this.options.type) === -1) {
        value = this._parseWriteEnum(value)
      }
    }
    return value
  }

  /**
   *
   * @param {number} value
   * @returns {string}
   * @private
   */
  _parseReadDate(value) {
    const d = value - 1
    const t = Math.round((d - Math.floor(d)) * 24 * 60 * 60)
    const date = new Date(1900, 0, d, 8, 0, t)

    const dp = `${date.getFullYear()}-${this._pad(date.getMonth() + 1)}-${this._pad(date.getDate())}`
    const tp = `${this._pad(date.getHours())}:${this._pad(date.getMinutes())}:${this._pad(date.getSeconds())}`

    return this.options.type === ColumnDef.types.DATETIME ? `${dp} ${tp}` : dp
  }

  _parseReadEnum(value) {
    for (const e of this.options.enums) {
      if (e.label === value) {
        return e.value
      }
    }
    throw new Error(`无效的枚举值 "${value}"`)
  }

  _parseWriteEnum(value) {
    for (const e of this.options.enums) {
      if (e.value === value) {
        return e.label
      }
    }
    throw new Error(`无效的枚举值 "${value}"`)
  }

  /**
   *
   * @param {string} value
   * @returns {string}
   * @private
   */
  _parseWriteDate(value) {
    return this.options.type === ColumnDef.types.DATETIME ? value : value.split(' ')[0]
  }

  _pad(num) {
    return num.toString().padStart(2, '0')
  }
}

/**
 * 列数据的类型
 * @type {{DATE: string, NUMBER: string, STRING: string}}
 *  如果表中有字段为 时间 类型，需要指定为 date 类型
 */
ColumnDef.types = {
  STRING: 'string',
  NUMBER: 'number',
  DATE: 'date',
  DATETIME: 'datetime'
}

/**
 * 表定义
 */
class SheetDef {
  /**
   *
   * @param {string} name Sheet名称
   * @param {ColumnDef[]} columns 列声明
   * @param {function({data: {}, index: number, raw: {}}): boolean | {}} [rowHandler] 行的值处理器。返回 false 表示值无效
   * @param {number} [maxRowCount] 最多读取的数据行数
   */
  constructor(name, columns, rowHandler, maxRowCount) {
    this.name = name

    /**
     *
     * @type {Map<string, ColumnDef>}
     */
    this.columns = new Map()
    this.rowHandler = rowHandler
    this.maxRowCount = maxRowCount

    columns.forEach(column => {
      this.columns.set(column.name, column)
    })
  }

  /**
   *
   * @param {WorkSheet} sheet
   */
  read(sheet) {
    const rows = XLSX.utils.sheet_to_json(sheet, {
      defval: null
    })

    const data = []

    for (let index = 0; index < rows.length; index++) {
      if (this.maxRowCount && index === this.maxRowCount) {
        break
      }

      const row = rows[index]
      const rowData = Object.create(null)

      // 先检查是否行的所有值都为空
      // 要是都为空，跳过此行
      if (Object.values(row).every(value => {
        return value === undefined || value === null || value === ''
      })) {
        continue
      }

      this.columns.forEach((column, name) => {
        if (!row.hasOwnProperty(name)) {
          throw new Error(`在表 "${this.name}" 中找不到列 "${name}"`)
        }
        try {
          rowData[column.field] = column.parseReadValue(row[name])
        } catch (e) {
          throw new Error(`表 "${this.name}" 第 ${index + 2} 行 "${name}" ${e.message}`)
        }
      })

      if (this.rowHandler) {
        const result = this.rowHandler({
          data: rowData,
          raw: row,
          index: index
        })
        if (result === false) {
          throw new Error(`表 "${this.name}" 第  ${index + 2} 行值无效`)
        }
        if (result !== undefined) {
          const type = /^\[object ([^[]+)]$/.exec(Object.prototype.toString.call(result))[1]
          if (type !== 'Object') {
            throw new Error(`表 "${this.name} 的行处理函数返回值类型 "${type}" 无效：仅支持返回 object/false 类型`)
          }
        }
      }

      data.push(rowData)
    }

    return data
  }

  /**
   * 将数据写入 Sheet
   * @param {object[]} data
   * @returns {WorkSheet}
   */
  write(data) {
    const header = []
    this.columns.forEach((col, colName) => {
      header.push(colName)
    })
    const rows = [header]
    data.forEach((row, index) => {
      const rowData = []
      this.columns.forEach(columnDef => {
        try {
          rowData.push(columnDef.parseWriteValue(row[columnDef.field]))
        } catch (e) {
          throw new Error(`表 "${this.name}" 的数据第 ${index} 行 "${name}" ${e.message}`)
        }
      })
      rows.push(rowData)
    })
    return XLSX.utils.aoa_to_sheet(rows)
  }
}

/**
 * 工作薄定义
 */
class WorkbookDef {
  /**
   *
   * @param {SheetDef[]} defs 此表格中要使用的表定义
   */
  constructor(defs) {
    /**
     *
     * @type {Map<string, SheetDef>}
     */
    this.sheets = new Map()

    defs.forEach(def => {
      this.sheets.set(def.name, def)
    })
  }

  /**
   * 读取文件内容
   * @param file
   * @returns {Promise<unknown>}
   */
  async readFile(file) {
    const reader = new FileReader()
    const promise = new Promise((resolve, reject) => {
      reader.onload = function () {
        resolve(reader.result)
      }
      reader.onerror = function (e) {
        reader.abort()
        reject(e)
      }
    })
    console.log(file)
    console.log(reader)
    reader.readAsArrayBuffer(file)

    return promise
  }

  /**
   * 从 文件对象 读取数据
   * @param {File} file
   */
  async read(file) {
    let dataBuffer = await this.readFile(file)

    const workbook = XLSX.read(dataBuffer, {
      type: 'array'
    })

    const result = []

    this.sheets.forEach((def, sheetName) => {
      if (workbook.SheetNames.indexOf(sheetName) === -1) {
        throw new Error(`找不到名称为 "${sheetName}" 的表`)
      }

      const sheet = workbook.Sheets[sheetName]

      result.push({
        name: sheetName,
        rows: def.read(sheet)
      })
    })

    return result
  }

  /**
   * 将数据写入文件
   * @param {[{name: string, rows: array}]} [data]
   * @param {string} saveAs 另存为文件名(会自动添加 xlsx 扩展名)
   */
  write(data, saveAs) {
    const workbook = {
      SheetNames: [],
      Sheets: Object.create(null)
    }

    const tempData = {}
    data.forEach(item => {
      tempData[item.name] = item.rows
    })

    this.sheets.forEach(sheetDef => {
      const name = sheetDef.name
      workbook.SheetNames.push(name)
      const rows = tempData[name] || []
      workbook.Sheets[name] = this.sheets.get(name).write(rows)
    })

    const output = XLSX.write(workbook, {
      bookType: 'xlsx', // 要生成的文件类型
      bookSST: false, // 是否生成Shared String Table，官方解释是，如果开启生成速度会下降，但在低版本IOS设备上有更好的兼容性
      type: 'binary'
    })

    const blobData = new Blob([string2arraybuffer(output)], {type: 'application/octet-stream'})

    if (!saveAs) {
      return blobData
    }

    downloadData(blobData, saveAs + '.xlsx')
  }
}

export {WorkbookDef, SheetDef, ColumnDef, ColumnEnum}
