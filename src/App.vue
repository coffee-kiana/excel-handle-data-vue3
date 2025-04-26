<script setup>
import * as XLSX from "xlsx";
import { read, writeFileXLSX } from "xlsx";
import { ref } from 'vue'
import { genFileId, } from 'element-plus'
const upload = ref()

const handleExceed = (files) => {
  upload.value.clearFiles()
  const file = files[0]
  file.uid = genFileId()
  upload.value.handleStart(file)
}

const excelData = ref([])
const newList2 = [];
const goodsObject = {};
const tableData = ref([]);
const secondSheetNewList = [];
const secondTableData = ref([]);
const secondOrderTable = ref([]);
const mergedTableData = ref([]);
const searchResultTable = ref([])



const displayTable1 = ref(true)
const displayTable2 = ref(true)
const displayTable3 = ref(true)

const handleFileChange = (uploadFile) => {
  const reader = new FileReader()
  reader.onload = (e) => {
    const data = new Uint8Array(e.target.result)
    const workbook = read(data, { type: 'array' })
    const firstSheet = workbook.Sheets[workbook.SheetNames[2]]
    const secondSheet = workbook.Sheets[workbook.SheetNames[0]]
    const goodsOrder2 = workbook.Sheets[workbook.SheetNames[1]]
    const guestToGoods = workbook.Sheets[workbook.SheetNames[3]]
    const guestToGoodsList =XLSX.utils.sheet_to_json(guestToGoods)
    const secondOrderJson = XLSX.utils.sheet_to_json(goodsOrder2)
    console.log('第二订单的数据:')
    const secondOrderList = secondOrderJson.map((obj, index) => {
      return {
        name: String(obj.Name).toUpperCase(),
        color: String(obj['Color & Size'].split('/')[0].trim()).toUpperCase(),
        styleNumber: obj.StyleNumber,
        sizeName: obj['Color & Size'].split('/')?.[1] ? String(obj['Color & Size'].split('/')[1]).replace(/\s+/g, '') : "",
        qty: obj.Quantity,
        tableRow: index + 2,
        isMerged: '',
        mergedIndex: 0,
      }
    })
    // console.log(secondOrderList)

    secondOrderTable.value = secondOrderList.sort((a, b) => {
      // 先按 name 排序
      if (a.name < b.name) return -1;
      if (a.name > b.name) return 1;
      // 如果 name 相同，则按 color 排序
      if (a.color < b.color) return -1;
      if (a.color > b.color) return 1;
    })

    const secondSheetJson = XLSX.utils.sheet_to_json(secondSheet)
    console.log('第二张表的数据:')
    // console.log(secondSheetJson)
    const secondSheetList = secondSheetJson.map(obj => {
      const newObj = {};
      for (const key in obj) {
        if (obj.hasOwnProperty(key)) {
          // 移除属性名中的空格
          const newKey = key.replace(/\s+/g, '');
          newObj[newKey] = obj[key];
        }
      }
      return newObj;
    });
    // console.log(secondSheetList)

    secondSheetList.forEach((item, index) => {
      for (const key in item) {
        if (key.includes('Size')) {
          secondSheetNewList.push({
            color: item.Color,
            name: item.Name,
            styleNumber: item.StyleNumber,
            price: item['M.S.R.P.(USD)'],
            sizeName: item[key].replace(/\s+/g, ''),
            qty: item[`${key.replace('Size', 'Qty')}`] || 0,
            tableRow: index + 2
          })

        }
      }
    })
    // console.log(secondSheetNewList)
    secondTableData.value = secondSheetNewList.sort((a, b) => {
      // 先按 name 排序
      if (a.name < b.name) return -1;
      if (a.name > b.name) return 1;

      // 如果 name 相同，则按 color 排序
      if (a.color < b.color) return -1;
      if (a.color > b.color) return 1;
    }).filter((item) => item.qty > 0)

    // 合并订单表
    // const mergedList = []
    const displayTableMerge = true

    let order1 = [...secondSheetNewList.filter(item => item.qty).sort((a, b) => {
      // 先按 name 排序
      if (a.name < b.name) return -1;
      if (a.name > b.name) return 1;
      // 如果 name 相同，则按 color 排序
      if (a.color < b.color) return -1;
      if (a.color > b.color) return 1;
    })]
    let order2 = [...secondOrderList]
    // console.log('合并前')
    // console.log(order1)
    // console.log(order2)

    let mergeList = order1.map((item1, indexOrder1) => {
      const findIndex = order2.findIndex((item2, indexOrder2) => {
        if (item1.name === item2.name && item1.color === item2.color && item1.sizeName === item2.sizeName) {
          return true
        }
      })
      if (findIndex !== -1) {
        order2[findIndex].isMerged = 'true'
        order2[findIndex].mergedIndex = indexOrder1 + 1
        return {
          ...item1,
          mergedQty: item1.qty + order2[findIndex].qty,
          isMerged: 'true',
          mergedIndex: findIndex + 1,
          original: '1&2'
        }
      } else {
        return {
          ...item1,
          mergedQty: item1.qty,
          isMerged: '',
          mergedIndex: '',
          original: '1'
        }
      }
    })
    order2.forEach((item, index) => {
      if (item.isMerged === '') {
        mergeList.push({
          ...item,
          mergedQty: item.qty,
          isMerged: '',
          mergedIndex: index + 1,
          original: '2'
        })
      }
    })
    mergedTableData.value = mergeList.sort((a, b) => {
      // 先按 name 排序
      if (a.name < b.name) return -1;
      if (a.name > b.name) return 1;
      // 如果 name 相同，则按 color 排序
      if (a.color < b.color) return -1;
      if (a.color > b.color) return 1;
    })

   


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
    // console.log(newList);

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
    // console.log(newList2);
    console.log('统计商品数量');
    // console.log(goodsObject);
    const guestList = []
    for (const key in goodsObject) {
      guestList.push({
        composeName: goodsObject[key].composeName,
        count: goodsObject[key].count,
        mapToExcel: goodsObject[key].mapToExcel,
      });
    }

    tableData.value = guestList.sort((a, b) => a.composeName.localeCompare(b.composeName));

    console.log('合并后')
    console.log(mergeList)
    // console.log('客户订单转商品维度:')
    // console.log(guestToGoodsList);
    
    const newGuestList = guestToGoodsList.map((item, index) => {
      const first=String(item.composeName.split("(")[0]).toUpperCase()
      const second=String(item.composeName.split("(")[1]).slice(0,-1)
      let temp={}
      if(second.includes(',')){
        temp.goodColor=String(second.split(',')[0]).toUpperCase(),
        temp.goodSize=String(second.split(',')[1]).toUpperCase()
      }else{
        temp.goodColor='',
        temp.goodSize=String(second.split(',')[0]).toUpperCase()
      }
      return {
        ...item,
        goodName: first,
        ...temp,
      }
    })
    console.log('拆分后的结果:')
    console.log(newGuestList)
    const searchResult=newGuestList.map((item,index)=>{
      const findIndex=mergeList.findIndex((obj)=>{
        if(obj.name===item.goodName &&obj.color.includes(item.goodColor)&&obj.sizeName===item.goodSize){
          return true
        }
      })
      if(findIndex!==-1){
        return {
          ...item,
          mergeQry:mergeList[findIndex].mergedQty,
          qtyCompare:mergeList[findIndex].mergedQty-item.count,
          mergeColor:mergeList[findIndex].color,
          mergeTableRow:findIndex+1,
          ...mergeList[findIndex]
        }
      }else{
        return {
          ...item,
          mergeQry:'',
          qtyCompare:'',
          mergeColor:'',
          mergeTableRow:''
        }
      }
    })
    console.log('查询结果:')
    console.log(searchResult)
    searchResultTable.value=searchResult.sort((a, b) => {
      // 先按 name 排序
      if (a.name < b.name) return -1;
      if (a.name > b.name) return 1;
      // 如果 name 相同，则按 color 排序
      if (a.color < b.color) return -1;
      if (a.color > b.color) return 1;
    })


  }
  reader.readAsArrayBuffer(uploadFile.raw)
}


</script>

<template>
  <div class="wrapper">
    <h2>读取excel</h2>

    <el-upload ref="upload" class="upload-demo" action="https://run.mocky.io/v3/9d059bf9-4660-45f2-925d-ce80ad6c4d15"
      :limit="1" :on-exceed="handleExceed" :auto-upload="false" :on-change="handleFileChange">
      <template #trigger>
        <el-button type="primary">select file</el-button>
      </template>
    </el-upload>



    <!-- <div style="margin-top: 20px;">
      <el-row style="margin-bottom: 16px;">
        <h3 style="margin-right: 20px;">订货单1按size统计</h3>
        <el-button type="primary" @click="displayTable2 = !displayTable2">显示/隐藏</el-button>
      </el-row>
      <div v-if="displayTable2">
        <el-table :data="secondTableData" border style="width: 100%">
          <el-table-column prop="name" label="商品名称" width="200" />
          <el-table-column prop="color" label="颜色规格" width="100" />
          <el-table-column prop="sizeName" label="尺码规格" />
          <el-table-column prop="qty" label="订货数量" width="200" />
          <el-table-column prop="tableRow" label="对应订货单的行" />
        </el-table>
      </div>
    </div>

    <div style="margin-top: 20px;">
      <el-row style="margin-bottom: 16px;">
        <h3 style="margin-right: 20px;">订货单2</h3>
        <el-button type="primary" @click="displayTable3 = !displayTable3">显示/隐藏</el-button>
      </el-row>
      <div v-if="displayTable3">
        <el-table :data="secondOrderTable" border style="width: 100%">
          <el-table-column prop="name" label="商品名称" width="200" />
          <el-table-column prop="color" label="颜色规格" width="100" />
          <el-table-column prop="sizeName" label="尺码规格" />
          <el-table-column prop="qty" label="订货数量" width="200" />
          <el-table-column prop="tableRow" label="对应订货单2的行" />
        </el-table>
      </div>
    </div> -->

    <div style="margin-top: 20px;">
      <el-row style="margin-bottom: 16px;">
        <h3 style="margin-right: 20px;">订单合并结果</h3>
        <el-button type="primary" @click="displayTable3 = !displayTable3">显示/隐藏</el-button>
      </el-row>
      <div v-if="displayTable3">
        <el-table :data="mergedTableData" border style="width: 100%">
          <el-table-column prop="name" label="商品名称" width="200" />
          <el-table-column prop="color" label="颜色规格" width="100" />
          <el-table-column prop="sizeName" label="尺码规格" />
          <el-table-column prop="qty" label="订货数量" width="200" />
          <el-table-column prop="tableRow" label="对应订货单的行" />
          <el-table-column prop="mergedQty" label="合并后数量" width="200" />
          <el-table-column prop="isMerged" label="是否有合并" />
          <el-table-column prop="mergedIndex" label="合并订货单2的行" />
          <el-table-column prop="original" label="来源订货单" />
        </el-table>
      </div>
    </div>

    <!-- <div style="margin-top: 20px;">
      <el-row style="margin-bottom: 16px;">
        <h3 style="margin-right: 20px;">客户订单按商品统计</h3>
        <el-button type="primary" @click="displayTable1 = !displayTable1">显示/隐藏</el-button>
      </el-row>
      <div v-if="displayTable1">
        <el-table :data="tableData" border style="width: 100%">
          <el-table-column prop="composeName" label="商品名称" width="200" />
          <el-table-column prop="count" label="下单总数" width="100" />
          <el-table-column prop="mapToExcel" label="对应客户订单Excel的行">
            <template #default="{ row }">
              <span v-for="ele in row.mapToExcel" :key="ele" class="tag-item" type="primary">
                {{ ele }},
              </span>
            </template>
          </el-table-column>
        </el-table>
      </div> 
    </div> -->

    <div style="margin-top: 20px;">
      <el-row style="margin-bottom: 16px;">
        <h3 style="margin-right: 20px;">客户订单查合并后订货单结果</h3>
        <el-button type="primary" @click="displayTable1 = !displayTable1">显示/隐藏</el-button>
      </el-row>
      <div v-if="displayTable1">
        <el-table :data="searchResultTable" border style="width: 100%">
          <el-table-column prop="composeName" label="下单商品名称" width="200" />
          <el-table-column prop="name" label="商品名称" width="200" />
          <el-table-column prop="count" label="下单总数" width="100" />
          <el-table-column prop="mapToExcel" label="对应客户订单Excel的行"/>
          <el-table-column prop="mergeQry" label="订货单总数" width="100" />
          <el-table-column prop="qtyCompare" label="订货-下单" width="100" />
          <el-table-column prop="mergeTableRow" label="对应合并后订货单的行" width="100" />
          <el-table-column prop="color" label="订货单颜色" width="100" />
          <el-table-column prop="sizeName" label="订货单尺寸名" width="100" />
        </el-table>
      </div>
    </div>

  </div>
</template>

<style scoped>
.wrapper {
  display: flex;
  flex-direction: column;
  width: 100%;
}
</style>
