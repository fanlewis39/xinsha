const async = require('async')
const officegen = require('officegen')
const fs = require('fs')
const http = require('http')
const url = require('url')
const mysql = require('mysql')
const path = require('path')
const DOMParser = require('xmldom').DOMParser

var sqlResult

// 清缓存，防止表格数据重复生成
function clearCacheAndRequire(url) {
  delete require.cache[require.resolve(url)]
  return file = require(url)
}

http.createServer(function (request, response) {
  var docx = officegen({
    type: 'docx',
    orientation: 'portrait',
    pageMargins: { top: 1000, left: 1000, bottom: 1000, right: 1000 }
  })

  docx.on('error', function (err) {
    console.log(err)
  })

  var data = url.parse(request.url,true).query

  //mysql的连接以及数据获取
  if( data.projectId && data.start && data.end ){
    var connection = mysql.createConnection({
      host: '192.168.0.15',
      user: "root",
      password: "123456",
      port: '3307',
      database: 'pm',
      multipleStatements: true
    })
    connection.connect()
    // 拉取周报任务数据初始化月报任务
    const sql1 = "SELECT C.*, D.ITEM_TOTALQUANTITY AS ITEM_本月累计 FROM ( SELECT A.APPLICATIONID, A.DOMAINID, A.ITEM_PID AS ITEM_任务ID, A.ITEM_NAME_ AS ITEM_任务名称, SUM(A.ITEM_PLANEDQUANTITY) AS ITEM_月内计划, SUM(A.ITEM_ACTUALQUANTITY) AS ITEM_月内实际, B.QUANTITIES_ AS ITEM_总工程量, B.UNIT_ AS ITEM_计量单位, B.WBS_ AS ITEM_任务序号, B.SUMMARY_ AS ITEM_摘要任务 FROM ( SELECT * FROM TLK_WEEKREPORT WHERE ITEM_PROJECTID = '" + data.projectId + "' AND ITEM_WEEK IN ( SELECT ITEM_WEEK AS ITEM_WEEK FROM TLK_WEEKLYREPORT_XS WHERE ITEM_PROJECTID = '" + data.projectId + "' AND ITEM_WEEKFINISH >= '" + data.start + "' AND ITEM_WEEKFINISH <= '" + data.end + "' )) A LEFT JOIN PLUS_TASK B ON A.ITEM_PID = B.PID GROUP BY A.ITEM_PID ) C LEFT JOIN ( SELECT ITEM_PID, SUM(ITEM_ACTUALQUANTITY) AS ITEM_TOTALQUANTITY FROM TLK_WEEKREPORT WHERE ITEM_PROJECTID = '" + data.projectId + "' AND ITEM_WEEK <= ( SELECT MAX(ITEM_WEEK) AS ITEM_WEEK FROM TLK_WEEKLYREPORT_XS WHERE ITEM_PROJECTID = '" + data.projectId + "' AND ITEM_WEEKFINISH >= '" + data.start + "' AND ITEM_WEEKFINISH <= '" + data.end + "' ) GROUP BY ITEM_PID ) D ON C.ITEM_任务ID = D.ITEM_PID order by C.ITEM_任务序号"
    // 拉取周报设施数据初始化月报任务
    const sql2 = ";SELECT ITEM_EQUIPMENT AS ITEM_设备名称,SUM(ITEM_QUANTITY) AS ITEM_数量,ITEM_规格 AS ITEM_型号规格,ITEM_UNIT AS ITEM_单位,'施工进行中' AS ITEM_状况描述 FROM TLK_EQUIPMENT WHERE ITEM_PROJECTID ='11e9-179e-d90bc6b5-96ef-1f01e45bff4e' AND ITEM_WEEK IN (SELECT ITEM_WEEK AS ITEM_WEEK FROM TLK_WEEKLYREPORT_XS WHERE ITEM_PROJECTID = '" + data.projectId + "' AND ITEM_WEEKFINISH >= '"+ data.start +"' AND ITEM_WEEKFINISH <= '" + data.end + "')GROUP BY ITEM_设备名称,ITEM_型号规格"
    // 拉取周报材料数据初始化月报任务
    const sql3 = ";SELECT ITEM_材料名称, ITEM_单位, ITEM_总量, SUM(ITEM_本周进场) AS ITEM_本月进场, MAX(ITEM_累计进场) AS ITEM_累计进场 FROM TLK_材料状况 WHERE ITEM_PROJECTID = '" + data.projectId + "' AND ITEM_WEEK IN ( SELECT ITEM_WEEK AS ITEM_WEEK FROM TLK_WEEKLYREPORT_XS WHERE ITEM_PROJECTID = '" + data.projectId + "' AND ITEM_WEEKFINISH >= '" + data.start + "' AND ITEM_WEEKFINISH <= '" + data.end + "' ) GROUP BY ITEM_材料名称"
    // 拉取周报检测数据初始化月报任务
    const sql4 = ";SELECT ITEM_检测内容 AS ITEM_检测内容, ITEM_单位, SUM(ITEM_本周检测 + 0) AS ITEM_本月检测, MAX(ITEM_检测累计情况 + 0) AS ITEM_累计检测, '合格' AS ITEM_结论 FROM TLK_检测情况 WHERE ITEM_PROJECTID = '"+ data.projectId +"' AND ITEM_WEEK IN ( SELECT ITEM_WEEK AS ITEM_WEEK FROM TLK_WEEKLYREPORT_XS WHERE ITEM_PROJECTID = '"+ data.projectId +"' AND ITEM_WEEKFINISH >= '"+ data.start +"' AND ITEM_WEEKFINISH <= '"+ data.end +"' ) GROUP BY ITEM_检测内容"
    // 拉取周报人员数据初始化月报任务
    const sql5 = ";SELECT ITEM_WORKER_CLASS AS ITEM_类别, SUM(ITEM_QUANTITY) AS ITEM_数量 FROM TLK_WORKER WHERE ITEM_PROJECTID = '" + data.projectId + "' AND ITEM_WEEK IN ( SELECT ITEM_WEEK AS ITEM_WEEK FROM TLK_WEEKLYREPORT_XS WHERE ITEM_PROJECTID = '" + data.projectId + "' AND ITEM_WEEKFINISH >= '" + data.start + "' AND ITEM_WEEKFINISH <= '" + data.end + "' ) GROUP BY ITEM_类别"
    // 拉取任务数据初始化月报下月计划
    const sql6 = ";SELECT ITEM_任务名称,ITEM_计量单位,ITEM_计划完成,ITEM_开始日期,ITEM_结束日期,ITEM_计划开始日期,ITEM_计划结束日期,ITEM_备注 FROM `tlk_月报下月计划` where item_所属项目 = '11e9-179e-d90bc6b5-96ef-1f01e45bff4e' and item_开始日期 = DATE_ADD('" + data.start + "',INTERVAL +1 MONTH) order by item_任务序号"
    const sql7 = ";select item_name,item_deptname from tlk_projectsetup where item_PROJECTID = '" + data.projectId + "'";
    // 文字部分
    const sql8 = ";select ITEM_开始日期,ITEM_结束日期,ITEM_期次,ITEM_工程质量状况,ITEM_文明施工情况,ITEM_文明施工计划,ITEM_上月问题落实,ITEM_本月存在问题 from tlk_月报主表 where ITEM_开始日期 = '" + data.start + "'"
    const sql = sql1 + sql2 + sql3 + sql4 + sql5 + sql6 + sql7 + sql8

    connection.query(sql, (err, result) => {
      if (err) {
        console.log('[SELECT ERROR]-', err.message)
        return;
      }
      sqlResult = []
      sqlResult = result
      docx_a()
    })
    connection.end()
  }

  //日期格式化
  Date.prototype.format = function(fmt) { 
    var o = {
       "M+" : this.getMonth()+1,                 //月份 
       "d+" : this.getDate(),                    //日 
       "h+" : this.getHours(),                   //小时 
       "m+" : this.getMinutes(),                 //分 
       "s+" : this.getSeconds(),                 //秒 
       "q+" : Math.floor((this.getMonth()+3)/3), //季度 
       "S"  : this.getMilliseconds()             //毫秒 
   }; 
   if(/(y+)/.test(fmt)) {
           fmt=fmt.replace(RegExp.$1, (this.getFullYear()+"").substr(4 - RegExp.$1.length))
   }
    for(var k in o) {
       if(new RegExp("("+ k +")").test(fmt)){
            fmt = fmt.replace(RegExp.$1, (RegExp.$1.length==1) ? (o[k]) : (("00"+ o[k]).substr((""+ o[k]).length)))
        }
    }
   return fmt
  }

  function docx_a() {
    const title = sqlResult[6][0].item_name
    const deptName = sqlResult[6][0].item_deptname
    const startYear = sqlResult[7][0].ITEM_开始日期.getFullYear()
    const startMonth = sqlResult[7][0].ITEM_开始日期.getMonth()+1
    const startDay = sqlResult[7][0].ITEM_开始日期.getDate();
    const finishYear = sqlResult[7][0].ITEM_结束日期.getFullYear()
    const finishMonth = sqlResult[7][0].ITEM_结束日期.getMonth()+1
    const finishDay = sqlResult[7][0].ITEM_结束日期.getDate()
    const months = sqlResult[7][0].ITEM_期次
    var pObj = docx.createP()

    // 首页
    pObj = docx.createP({ align: 'center' })
    pObj = docx.createP()
    pObj.addText(title, { font_size: 17 })
    pObj = docx.createP()
    pObj = docx.createP()
    pObj = docx.createP()
    pObj = docx.createP()
    pObj = docx.createP({ align: 'center' })
    pObj.addText('施工月报', { font_size: 35 })
    pObj = docx.createP()
    pObj = docx.createP({ align: 'center' })
    pObj.addText('（第  ', { font_size: 17 })
    pObj.addText("0" + months, { bold: true, underline: true, font_size: 17 })
    pObj.addText('  期，编号：', { font_size: 17 })
    pObj.addText('0' + months, { bold: true, underline: true, font_size: 17 })
    pObj.addText('）', { font_size: 17 })
    pObj = docx.createP({ align: 'center' })
    pObj.addText('（', { font_size: 17 })
    pObj.addText('' + startYear, { bold: true, underline: true, font_size: 17 })
    pObj.addText('年', { font_size: 17 })
    pObj.addText('' + startMonth, { bold: true, underline: true, font_size: 17 })
    pObj.addText('月', { font_size: 17 })
    pObj.addText('' + startDay, { bold: true, underline: true, font_size: 17 })
    pObj.addText('日至', { font_size: 17 })
    pObj.addText('' + finishYear, { bold: true, underline: true, font_size: 17 })
    pObj.addText('年', { font_size: 17 })
    pObj.addText('' + finishMonth, { bold: true, underline: true, font_size: 17 })
    pObj.addText('月', { font_size: 17 })
    pObj.addText('' + finishDay, { bold: true, underline: true, font_size: 17 })
    pObj.addText('日）', { font_size: 17 })
    pObj = docx.createP()
    pObj = docx.createP()
    pObj = docx.createP({ align: 'center' })
    pObj.addText('编写：_________', { font_size: 17 })
    pObj = docx.createP({ align: 'center' })
    pObj.addText('审核：_________', { font_size: 17 })
    pObj = docx.createP()
    pObj = docx.createP()
    pObj = docx.createP()
    pObj = docx.createP({ align: 'center' })
    pObj.addText('中交第四航务工程局有限公司' + title, { font_size: 17 })
    pObj.addText(deptName, { font_size: 17 })
    pObj = docx.createP({ align: 'center' })
    pObj.addText('编制日期：' + finishYear + '年' +finishMonth + '月' + finishDay + '日', { font_size: 17 })
    pObj = docx.createP()
    pObj = docx.createP()

    // 编号表格
    var tab1_no = '0' + months;
    var tab1_time = startYear + '-' + startMonth + '-' + startDay + '~' + finishYear + '-' + finishMonth + '-' + finishDay;
    var tab1_unit = '中交第四航务工程局有限公司' + title + deptName;
    var tab1_publish = finishYear + '-' + finishMonth + '-' + finishDay;
    var table1 = [
      [{
        val: '编号',
        opts: {
          cellColWidth: 1000,
          // color: '000000',
          sz: 30,
          shd: {
              fill: '000000',
              themeFill: 'text1'
          }
        }
      },{
        val: tab1_no,
        opts: {
          cellColWidth: 2500,
          // color: '000000',
          sz: 30,
          shd: {
              fill: '000000',
              themeFill: 'text1'
          }
        }
      },{
        val: '时间段',
        opts: {
          cellColWidth: 1000,
          // color: '000000',
          sz: 30,
          shd: {
              // fill: '000000',
              themeFill: 'text1'
          }
        }
      },{
        val: tab1_time,
        opts: {
          cellColWidth: 2500,
          color: '000000',
          sz: 30,
          shd: {
              fill: '000000',
              themeFill: 'text1'
          }
        }
      }],
      ['主送单位','广州港工程管理有限公司','发文单位',tab1_unit],
      ['发文日期',tab1_publish,'发文人','胡少文'],
    ]

    // 分别引入各表格模块
    // 本月工程进度状况
    var table2 = clearCacheAndRequire('./regularTable/table2')
    for( let i in sqlResult[0] ){
      var sqlData = sqlResult[0][i]
      var percent = 0;
      if( sqlData.ITEM_总工程量 != 0 ){
        percent = (100*(sqlData.ITEM_本月累计/sqlData.ITEM_总工程量)).toFixed(1)
      }
      var arr = [sqlData.ITEM_任务序号,sqlData.ITEM_任务名称,sqlData.ITEM_计量单位,sqlData.ITEM_月内计划,sqlData.ITEM_月内实际,sqlData.ITEM_本月累计 + "/" + sqlData.ITEM_总工程量,percent]
      table2.push(arr)
    }

    // 设备情况
    var table3 = clearCacheAndRequire('./regularTable/table3')
    var num = 1;
    for( let i in sqlResult[1] ){
      var sqlData = sqlResult[1][i]
      var arr = [num++,sqlData.ITEM_设备名称,sqlData.ITEM_型号规格,sqlData.ITEM_单位,sqlData.ITEM_数量,sqlData.ITEM_状况描述];
      table3.push(arr)
    }

    // 材料进场情况
    var table4 = clearCacheAndRequire('./regularTable/table4')
    var num = 1;
    for( let i in sqlResult[2] ){
      var sqlData = sqlResult[2][i]
      var percent = (100*(sqlData.ITEM_累计进场/sqlData.ITEM_总量)).toFixed(1)
      var arr = [num++,sqlData.ITEM_材料名称,sqlData.ITEM_单位,sqlData.ITEM_总量,sqlData.ITEM_本月进场,sqlData.ITEM_累计进场,percent]
      table4.push(arr)
    }

    // 检测情况
    var table5 = clearCacheAndRequire('./regularTable/table5')
    var num = 1
    for( let i in sqlResult[3] ){
      var sqlData = sqlResult[3][i]
      var arr = [num++,sqlData.ITEM_检测内容,sqlData.ITEM_单位,sqlData.ITEM_本月检测,sqlData.ITEM_累计检测,sqlData.ITEM_结论]
      table5.push(arr)
    }

    // 本月人员配置状况
    var table6 = clearCacheAndRequire('./regularTable/table6')
    var num = 1
    var sum = 0
    for( let i in sqlResult[4] ){
      var sqlData = sqlResult[4][i]
      if( sqlData.ITEM_类别 != '合计' ){
        sum += (sqlData.ITEM_数量*1);
        var arr = [num++,sqlData.ITEM_类别,sqlData.ITEM_数量,'已进场开展工作']   
        table6.push(arr)    
      }
    }
    table6.push([num++,'合计',sum,'已进场开展工作'])

    // 下月施工计划
    var table7 = clearCacheAndRequire('./regularTable/table7')
    var num = 1
    for( let i in sqlResult[5] ){
      var sqlData = sqlResult[5][i]
      var start = sqlData.ITEM_开始日期.format("yyyy-MM-dd")
      var end = sqlData.ITEM_结束日期.format("yyyy-MM-dd")
      if( sqlData.备注 == null ) sqlData.备注 = ""
      var arr = [num++,sqlData.ITEM_任务名称,sqlData.ITEM_计量单位,sqlData.ITEM_计划完成,start+'~'+end,sqlData.ITEM_备注]
      table7.push(arr)
    }

    // 文档文字部分
    const engineeringQuality = sqlResult[7][0].ITEM_工程质量状况.replace(/<[^<>]+>/g,"")
    // 施工情况
    const civilizedConstruction = sqlResult[7][0].ITEM_文明施工情况
    // 施工计划
    const civilizedConstructionPlan = sqlResult[7][0].ITEM_文明施工计划
    // 上月问题落实
    const implementationStatus = sqlResult[7][0].ITEM_上月问题落实
    // 本月存在问题
    const problem = sqlResult[7][0].ITEM_本月存在问题

    //施工情况排版处理
    function civilizedConstruction_deal(str) {
      var doc = new DOMParser().parseFromString(str)
      var childNodes = doc.childNodes
      console.log(childNodes)
      if( childNodes[0].textContent.indexOf("   ") != 0 ){
        data2.push({
          type: "text",
          val: childNodes[0].textContent,
          opt:{ font_size: 14, align: 'center' }
        })
      }
      for( var i=0;i<childNodes.length;i++ ){
        // if(childNodes[i].childNodes.length == 1 && childNodes[i].textContent.indexOf("   ") != 0 )
        // if( childNodes[i].tagName == 'table' ) {
        //   console.log("true")
        // }
        var grandChildNodes = childNodes[i].childNodes
        for( var j=0;j<grandChildNodes.length;j++ ){
          if( grandChildNodes[j].tagName == 'img' ){
            data2.push({
                type: "image",
                path: path.resolve( "../../../../.." + grandChildNodes[j].getAttribute("src")),
                opt: { cx: 400, cy: 300 }
            })
          }else{
            if( grandChildNodes[j].childNodes != null ){
              var dGrandChildNodes = grandChildNodes[j].childNodes
              for( var k=0;k<dGrandChildNodes.length;k++ ){
                if( dGrandChildNodes[k].tagName == 'img' ){
                  data2.push({
                    type: "image",
                    path: path.resolve( "../../../../.." + dGrandChildNodes[k].getAttribute("src")),
                    opt: { cx: 400, cy: 300 }
                  })
                }
              }
            }           
          }
        }
        if( (i+1)<childNodes.length ){
          data2.push({
              type: "text",
              val: childNodes[i+1].textContent,
              opt:{ font_size: 14, align: 'center' }
          })
        }
      }
    }

    // 施工计划排版处理
    function civilizedConstructionPlan_deal(str) {
      var doc = new DOMParser().parseFromString(str)
      var childNodes = doc.childNodes
      data2.push({
          type: "text",
          val: "2、下月安全文明施工管理计划",
          opt: { font_size: 15 }
      })
      for( var i=0;i<childNodes.length;i++ ){
        if( childNodes[i].textContent != "" ){
          data2.push({
              type: "text",
              val: childNodes[i].textContent,
              opt: { font_size: 15 }
          })
        }
      }
    }

    // 上月问题落实
    function implementationStatus_deal(str) {
      var doc = new DOMParser().parseFromString(str)
      var childNodes = doc.childNodes
      data2.push({
        type: "text",
        val: "六、上月施工中遇到的问题及落实情况",
        opt: { bold: true, font_size: 17 }
      })
      for( var i=0;i<childNodes.length;i++ ){
        if( childNodes[i].textContent != "" ){
          data2.push({
              type: "text",
              val: childNodes[i].textContent,
              opt: { font_size: 15 }
          })
        }
      }
    }

    // 本月存在问题
    function problem_deal(str) {
      var doc = new DOMParser().parseFromString(str)
      var childNodes = doc.childNodes
      data2.push({
        type: "text",
        val: "七、本月存在的问题及解决措施",
        opt: { bold: true, font_size: 17 }
      })
      for( var i=0;i<childNodes.length;i++ ){
        if( childNodes[i].textContent != "" ){
          data2.push({
              type: "text",
              val: childNodes[i].textContent,
              opt: { font_size: 15 }
          })
        }
      }
    }
    // 调试：http://localhost:4005/?projectId=11e9-179e-d90bc6b5-96ef-1f01e45bff4e&start=2018-07-26%2000:00:00&end=2018-08-25%2023:59:59

    var tableStyle = {
      tableColWidth: 6000,
      tableSize: 30,
      borders: true,
      borderSize: 10,
      color: '000000',
      tableAlign: 'center',
      tableFontFamily: 'Times New Roman'
    }

    var data1 = [
      {
        type: "table",
        val: table1,
        opt: tableStyle
      }, {
        type: 'createp'
      }, {
        type: "text",
        val: "一、本月工程进度状况",
        opt: { bold: true, font_size: 17 }
      }, {
        type: "table",
        val: table2,
        opt: tableStyle
      }, {
        type: 'createp'
      }, {
        type: "text",
        val: "二、本月工程质量状况",
        opt: { bold: true, font_size: 17 }
      }, {
        type: "text",
        val: engineeringQuality,
        opt: { font_size: 15 }
      }, {
        type: 'createp'
      }, {
        type: "text",
        val: "三、本月设备/材料状况",
        opt: { bold: true, font_size: 17 }
      }, {
        type: "text",
        val: "1、设备情况",
        opt: { font_size: 15 }
      }, {
        type: "table",
        val: table3,
        opt: tableStyle
      }, {
        type: 'createp'
      }, {
        type: "text",
        val: "2、材料进场情况",
        opt: { font_size: 15 }
      }, {
        type: "table",
        val: table4,
        opt: tableStyle
      }, {
        type: 'createp'
      }, {
        type: "text",
        val: "3、检测情况",
        opt: { font_size: 15 }
      }, {
        type: "table",
        val: table5,
        opt: tableStyle
      }, {
        type: 'createp'
      }, {
        type: "text",
        val: "四、本月人员配置状况",
        opt: { bold: true, font_size: 17 }
      }, {
        type: "table",
        val: table6,
        opt: tableStyle
      }, {
        type: 'createp'
      }, {
        type: "text",
        val: "五、HSE方面",
        opt: { bold: true, font_size: 17 }
      }, {
        type: "text",
        val: "1、本月安全生产、文明施工情况",
        opt: { font_size: 15 }
      }]

      var data2 = []
      civilizedConstruction_deal(civilizedConstruction.replace(/&nbsp;/g,""))
      civilizedConstructionPlan_deal(civilizedConstructionPlan)
      implementationStatus_deal(implementationStatus)
      problem_deal(problem)

      var data3 = [{
        type: 'createp'
      }, {
        type: "text",
        val: "八、下月施工计划",
        opt: { bold: true, font_size: 17 }
      }, {
        type: "table",
        val: table7,
        opt: tableStyle
      }
    ]
    
    docx.createByJson(data1)
    docx.createByJson(data2)
    docx.createByJson(data3)

    //header
    var header = docx.getHeader().createP();
    header.addText('    中交第四航务工程局有限公司');
    header.addHorizontalLine()

    var fileName = 'monthreport0' + months;

    //头页
    pObj = docx.createP({ align: 'center' })

    pObj = docx.createP()
    var out = fs.createWriteStream('example.docx')
    out.on('error', function (err) {
      console.log(err)
    })
    async.parallel([
      function (done) {
        out.on('close', function () {
          console.log('Finish to create a DOCX file.')
          done(null)
        })
        docx.generate(out)
      }
    ], function (err) {
      if (err) {
        console.log('error: ' + err)
      } // Endif.
    })
    docx.generate(response)

    response.writeHead(200, {
      'Access-Control-Allow-Origin': '*',
      'Content-Type': 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
      'Content-disposition': 'attachment; filename=' + fileName + '.docx',
      'Access-Control-Allow-Credentials': 'true',
      'Access-Control-Allow-Methods': 'PUT,POST,GET,DELETE,OPTIONS'
    });
  }
}).listen(4005)