<template>
  <div class="container">
    <div style="display:flex;">
      <el-upload
        ref="upload"
        action="action"
        accept="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        :show-file-list="false"
        :on-change="onUploadChange"
      >
        <el-button
          size="small"
          type="primary"
        >点击上传</el-button>
      </el-upload>

      <el-button
        size="small"
        @click="handleDownload"
      >点击下载</el-button>

    </div>

    <el-table
      :data="tableData"
      border
      style="width: 100%"
    >
      <el-table-column v-for="key in Object.keys(tableData[0]|| {})"
        :prop="key"
        :label="key"
      >
      </el-table-column>
    </el-table>
  </div>
</template>

<script>
import { readExcelToJson, saveJsonToExcel } from './utils.js';

export default {
  data() {
    return {
      file: null,
      tableData: [],
      outputName:"output.xlsx"
    };
  },

  methods: {
    // 读取文件为json数据
    onUploadChange(file) {
      console.log(file);
      this.file = file;
      readExcelToJson(file).then((res) => {
        this.tableData = res;
        console.log("res",res)
      });
    },

    handleDownload() {
      this.addData()
      
    },
    addData(){
    this.$axios.get("/api/data/sk/101110101.html").then(res=>{
      let newTableData=JSON.parse(JSON.stringify(this.tableData))
           let lastData=newTableData.map(item=>{
             return{
               "地区编码":101110101,
               "地区名称":item["地区名称"],
               "天气":res.data.weatherinfo.temp
             }
           })
           
           saveJsonToExcel(lastData,this.outputName);
    })
    },
  },
};
</script>

<style>
body {
  background: #f4f4f4;
  padding: 0;
  margin: 0;
}
.container {
  width: 1024px;
  min-height: 100vh;
  margin: 0 auto;
  padding: 20px;
  background: #fff;
}
</style>
