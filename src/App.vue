<script setup>
import * as XLSX from "xlsx";
import { read, writeFileXLSX } from "xlsx";
import { ref } from 'vue'
import { genFileId, } from 'element-plus'
const upload = ref()

const handleExceed= (files) => {
  upload.value.clearFiles()
  const file = files[0]
  file.uid = genFileId()
  upload.value.handleStart(file)
}

const excelData = ref([])
const newList2 = [];
const goodsObject = {};
const tableData = ref([]);
const handleFileChange = (uploadFile) => {
  const reader = new FileReader()
  reader.onload = (e) => {
    const data = new Uint8Array(e.target.result)
    const workbook = read(data, { type: 'array' })
    const firstSheet = workbook.Sheets[workbook.SheetNames[0]]
    const secondSheet = workbook.Sheets[workbook.SheetNames[1]]
    const secondSheetJson = XLSX.utils.sheet_to_json(secondSheet)
    console.log('第二张表的数据:')
    console.log(secondSheetJson)
    const json = XLSX.utils.sheet_to_json(firstSheet)
    excelData.value = json
    console.log('解析后的数据:', excelData.value)
    // 更换属性名称
    const newList = json.map((item) => {
          return {
            userName: item["微信昵称"],
            countPrice: item["订单总金额"],
            wechatSeqId: item["接龙号"],
            goodList: item["商品合计"].split("\n").filter(Boolean),
          };
        });
        console.log(newList);
        
        // 按商品维度拆分数据
        newList.forEach((item, rowIndex) => {
          item.goodList.forEach((good, index) => {
            newList2.push({
              userName: item.userName,
              countPrice: item.countPrice,
              wechatSeqId: item.wechatSeqId,
              rowId: rowIndex + 2,
              userGoodsIndex: index + 1,
              goodName: good.split("*")[0].trim(),
              goodCount: good.split("*")[1].trim(),
            });
            goodsObject[good.split("*")[0].trim()] = {
              composeName: good.split("*")[0].trim(),
              count: 0,
              mapToExcel: [],
            };
          });
        });
        console.log('按商品维度拆分数据');

        // 统计商品数量
        newList2.forEach((item) => {
          goodsObject[item.goodName].count += parseInt(item.goodCount);
          goodsObject[item.goodName].mapToExcel = [
            ...goodsObject[item.goodName].mapToExcel,
            `${item.rowId}-${item.userGoodsIndex}`,
          ];
        });
        console.log(newList2);
        console.log('统计商品数量');
        console.log(goodsObject);

        for (const key in goodsObject) {
          tableData.value.push({
            composeName: goodsObject[key].composeName,
            count: goodsObject[key].count,
            mapToExcel: goodsObject[key].mapToExcel,
          });
        }

        tableData.value.sort((a, b) => a.composeName.localeCompare(b.composeName));
        
  }
  reader.readAsArrayBuffer(uploadFile.raw)
}


</script>

<template>
  <div class="wrapper">
  <h2>读取excel</h2>

  <el-upload
    ref="upload"
    class="upload-demo"
    action="https://run.mocky.io/v3/9d059bf9-4660-45f2-925d-ce80ad6c4d15"
    :limit="1"
    :on-exceed="handleExceed"
    :auto-upload="false"
    :on-change="handleFileChange" 
  >
    <template #trigger>
      <el-button type="primary">select file</el-button>
    </template>
    <template #tip>
      <div class="el-upload__tip text-red">
        limit 1 file, new file will cover the old file
      </div>
    </template>
  </el-upload>

  <div style="margin-top: 20px;">
    <el-table :data="tableData" border style="width: 100%">
    <el-table-column prop="composeName" label="商品名称" width="200" />
    <el-table-column prop="count" label="下单总数" width="100" />
    <el-table-column prop="mapToExcel" label="对应客户订单Excel的行" >
      <template #default="{ row }">
        <span
          v-for="ele in row.mapToExcel"
          :key="ele"
          class="tag-item"
          type="primary"
        >
          {{ ele }},
        </span>
      </template>
    </el-table-column>
  </el-table>
  </div>
</div>
</template>

<style scoped>
.wrapper{
  display: flex;
  flex-direction: column;
}
</style>
