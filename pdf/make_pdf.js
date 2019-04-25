const fs = require("fs");
const http = require("http");
const moment = require("moment");
const html = fs.readFileSync("./template.html", "utf-8");
const mysql = require("mysql");
const url = require("url");
const express = require("express");
const app = express();
const create = require("./create_pdf.js");

var sqlResult = [];

app.get('/', function (req, res) {
  const docid = req.query.docid;
  const projectid = req.query.projectid;
  //连接mysql
  const connection = mysql.createConnection({
    host:'localhost',
    user:'root',
    password:'123456',
    port:'3307',
    database:'pm',
    multipleStatements: true
  });
  connection.connect();
  const sql1 = "select item_week,date_format(item_weekstart,'%Y-%m-%d') as item_weekstart ,date_format(item_weekstart,'%Y') as item_weekstartYear ,date_format(item_weekstart,'%m') as item_weekstartMonth,date_format(item_weekstart,'%d') as item_weekstartDay, date_format(item_weekfinish,'%Y-%m-%d') as item_weekfinish ,date_format(item_weekfinish,'%Y') as item_weekfinishYear ,date_format(item_weekfinish,'%m') as item_weekfinishMonth,date_format(item_weekfinish,'%d') as item_weekfinishDay from tlk_weeklyreport_xs where id='" + docid + "'";
  const sql2 = ";select item_name,item_deptname from tlk_projectsetup where item_PROJECTID = '" + projectid + "'";
  const renwujindu_sql = ";select item_search_,item_name_,item_planedquantity,ITEM_ACTUALQUANTITY,ITEM_THEASSIGNMENT,ITEM_PLANQUANTITY,ITEM_NOTES_,ITEM_RECTIFICATION,FORMAT(ITEM_PERCENTCOMPLETE_,2) as ITEM_PERCENTCOMPLETE from tlk_weekreport where parent = '" + docid + "' and (ITEM_SUMMARY_ = 1 or ITEM_ACTUALQUANTITY > 0 or (ITEM_THEASSIGNMENT <> '' and ITEM_THEASSIGNMENT is not null )) order by item_search_ asc";
  const yanchikaigong_sql = ";select item_search_,item_name_,item_planedquantity,ITEM_ACTUALQUANTITY,ITEM_THEASSIGNMENT,ITEM_PLANQUANTITY,ITEM_NOTES_,ITEM_RECTIFICATION,FORMAT(ITEM_PERCENTCOMPLETE_,2) as ITEM_PERCENTCOMPLETE from tlk_weekreport where parent = '" + docid + "' and ITEM_SUMMARY_ = 0 and ITEM_ACTUALQUANTITY = 0 and (ITEM_THEASSIGNMENT = '' or ITEM_THEASSIGNMENT is null ) order by item_search_ asc";
  const shebeizhuangkuang_sql = ";select ITEM_EQUIPMENT,`ITEM_规格`,FORMAT(ITEM_QUANTITY,0) as ITEM_QUANTITY,ITEM_TOTAL,`ITEM_状态` from tlk_equipment where parent = '" + docid + "'";
  const cailiaozhuangkuang_sql = ";select `ITEM_材料名称`,`ITEM_单位`,`ITEM_总量`,`ITEM_本周进场`,`ITEM_累计进场`,`ITEM_累计进场比例` from `tlk_材料状况` where parent = '" + docid + "'";
  const jianceqingkuang_sql = ";select `ITEM_检测内容`,`ITEM_单位`,`ITEM_本周检测`,`ITEM_检测累计情况`,`ITEM_结论` from `tlk_检测情况` where parent = '" + docid + "'";
  const renyuanpeizhi_sql = ";select ITEM_WORKER_CLASS,ITEM_QUANTITY,`ITEM_备注` from tlk_worker where parent = '" + docid + "'";
  const anquanzhiliangneiyye_sql = ";select `ITEM_安全`,`ITEM_质量`,`ITEM_联系单`,`ITEM_问题落实`,`ITEM_影响协商`,item_照片,`ITEM_沉降位移` from tlk_weeklyreport_xs where id='" + docid + "'";
  const sql = sql1 + sql2 + renwujindu_sql + yanchikaigong_sql + shebeizhuangkuang_sql + cailiaozhuangkuang_sql + jianceqingkuang_sql + renyuanpeizhi_sql + anquanzhiliangneiyye_sql;
  connection.query(sql,function(err,result){
    if(err){
      console.log("err--",err.message);
      return;
    }  
    sqlResult = result;
    HtmlToPdf();
  });
  connection.end();

  function HtmlToPdf(){
    var options = {
      format: "A4",
      paginationOffset: 4,
      header: {
        height: "30mm",
        contents: '<div><img src="file:///C:/cnqisoft/bin/apache-tomcat-7.0.57/webapps/Z0162S/portal/H5/officegen/imagelogo.png"></div><br><span>中交第四航务工程局有限公司</span><br><span style="font-size: 12px">CCCC Fourth harbor Engineering Co.,Lrd.</span></br>_______________________________________________________________________________________'
      },
      footer: {
        height: "20mm",
        paginationOffset: -1,
        contents:{
          1: " ",
          default: '<center><span style="color: #444;font-size: 13px">' + sqlResult[1][0].item_name + '水工土建工程&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp施工周报—{{page}}—</span></center>'
        }
      },
      border: {
        left: "10mm",
        right: "10mm"
      }
    };
    //第一页数据
    // console.log(sqlResult)
    const title = sqlResult[1][0].item_name;
    const week = sqlResult[0][0].item_week;
    const weekstartYear = sqlResult[0][0].item_weekstartYear;
    const weekstartMonth = sqlResult[0][0].item_weekstartMonth;
    const weekstartDay = sqlResult[0][0].item_weekstartDay;
    const weekfinishYear = sqlResult[0][0].item_weekfinishYear;
    const weekfinishMonth = sqlResult[0][0].item_weekfinishMonth;
    const weekfinishDay = sqlResult[0][0].item_weekfinishDay;
    const deptname = sqlResult[1][0].item_deptname;
    //表格数据
    const table_arr = [];
    table_arr[0] = [];table_arr[1] = [];table_arr[2] = [];table_arr[3] = [];table_arr[4] = [];table_arr[5] = [];
    for( let j=2;j<8;j++ ){
      for( let i=0;i<sqlResult[j].length;i++ ){
        table_arr[j-2].push(sqlResult[j][i]);
      }
    }
    var table1 = "<table border='1'><tr><td style='width:45px'>WBS</td><td>名称</td><td style='width:45px'>本周计划</td><td style='width:45px'>实际完成</td><td style='width:85px'>作业情况</td><td style='width:50px'>下周计划</td><td style='width:75px'>偏离说明</td><td style='width:75px'>纠偏措施</td><td style='width:45px'>完成率</td></tr>";
    if( table_arr[0].length < 21 ){
      for( let i=0;i<table_arr[0].length;i++){
        if( table_arr[0][i].ITEM_THEASSIGNMENT == '1' ){
          table_arr[0][i].ITEM_THEASSIGNMENT = '进场准备'
        }else if( table_arr[0][i].ITEM_THEASSIGNMENT == '2' ){
          table_arr[0][i].ITEM_THEASSIGNMENT = '计划性停工'
        }else if( table_arr[0][i].ITEM_THEASSIGNMENT == '3' ){
          table_arr[0][i].ITEM_THEASSIGNMENT = '意外性停工'
        }
        if( table_arr[0][i].ITEM_NOTES_ == null ){ table_arr[0][i].ITEM_NOTES_ = "" }
        if( table_arr[0][i].ITEM_RECTIFICATION == null ){ table_arr[0][i].ITEM_RECTIFICATION = "" }
        table1 += "<tr><td>"+ table_arr[0][i].item_search_ + "</td><td>"+ table_arr[0][i].item_name_ + "</td><td>" + table_arr[0][i].item_planedquantity + "</td><td>"+ table_arr[0][i].ITEM_ACTUALQUANTITY + "</td><td>"+ table_arr[0][i].ITEM_THEASSIGNMENT + "</td><td>"+ table_arr[0][i].ITEM_PLANQUANTITY + "</td><td>"+ table_arr[0][i].ITEM_NOTES_ + "</td><td>" + table_arr[0][i].ITEM_RECTIFICATION + "</td><td>"+ table_arr[0][i].ITEM_PERCENTCOMPLETE +"</td></tr>";
      }
      table1 += "</table>";
    }else{
      for( let i=0;i<21;i++){
        if( table_arr[0][i].ITEM_THEASSIGNMENT == '1' ){
          table_arr[0][i].ITEM_THEASSIGNMENT = '进场准备'
        }else if( table_arr[0][i].ITEM_THEASSIGNMENT == '2' ){
          table_arr[0][i].ITEM_THEASSIGNMENT = '计划性停工'
        }else if( table_arr[0][i].ITEM_THEASSIGNMENT == '3' ){
          table_arr[0][i].ITEM_THEASSIGNMENT = '意外性停工'
        }
        if( table_arr[0][i].ITEM_NOTES_ == null ){ table_arr[0][i].ITEM_NOTES_ = "" }
        if( table_arr[0][i].ITEM_RECTIFICATION == null ){ table_arr[0][i].ITEM_RECTIFICATION = "" }
        table1 += "<tr><td>"+ table_arr[0][i].item_search_ + "</td><td>"+ table_arr[0][i].item_name_ + "</td><td>" + table_arr[0][i].item_planedquantity + "</td><td>"+ table_arr[0][i].ITEM_ACTUALQUANTITY + "</td><td>"+ table_arr[0][i].ITEM_THEASSIGNMENT + "</td><td>"+ table_arr[0][i].ITEM_PLANQUANTITY + "</td><td>"+ table_arr[0][i].ITEM_NOTES_ + "</td><td>" + table_arr[0][i].ITEM_RECTIFICATION + "</td><td>"+ table_arr[0][i].ITEM_PERCENTCOMPLETE +"</td></tr>";
      }
      table1 += "</table><table border='1' style='margin-top:500px'>";
      for( let i=21;i<table_arr[0].length;i++){
        if( table_arr[0][i].ITEM_THEASSIGNMENT == '1' ){
          table_arr[0][i].ITEM_THEASSIGNMENT = '进场准备'
        }else if( table_arr[0][i].ITEM_THEASSIGNMENT == '2' ){
          table_arr[0][i].ITEM_THEASSIGNMENT = '计划性停工'
        }else if( table_arr[0][i].ITEM_THEASSIGNMENT == '3' ){
          table_arr[0][i].ITEM_THEASSIGNMENT = '意外性停工'
        }
        if( table_arr[0][i].ITEM_NOTES_ == null ){ table_arr[0][i].ITEM_NOTES_ = "" }
        if( table_arr[0][i].ITEM_RECTIFICATION == null ){ table_arr[0][i].ITEM_RECTIFICATION = "" }
        table1 += "<tr><td style='width:45px'>"+ table_arr[0][i].item_search_ + "</td><td>"+ table_arr[0][i].item_name_ + "</td><td style='width:45px'>" + table_arr[0][i].item_planedquantity + "</td><td style='width:45px'>"+ table_arr[0][i].ITEM_ACTUALQUANTITY + "</td><td style='width:85px'>"+ table_arr[0][i].ITEM_THEASSIGNMENT + "</td><td style='width:50px'>"+ table_arr[0][i].ITEM_PLANQUANTITY + "</td><td style='width:75px'>"+ table_arr[0][i].ITEM_NOTES_ + "</td><td style='width:75px'>" + table_arr[0][i].ITEM_RECTIFICATION + "</td><td style='width:45px'>"+ table_arr[0][i].ITEM_PERCENTCOMPLETE +"</td></tr>";
      }
      table1 += "</table>";
    }

    var table1p = "<table border='1'><tr><td>WBS</td><td>名称</td><td>本周计划</td><td>实际完成</td><td>下周计划</td><td>偏离说明</td><td>纠偏措施</td></tr>";
    for( let i=0;i<table_arr[1].length;i++){
      if( table_arr[1][i].ITEM_NOTES_ == null ){ table_arr[1][i].ITEM_NOTES_ = "" }
      if( table_arr[1][i].ITEM_RECTIFICATION == null ){ table_arr[1][i].ITEM_RECTIFICATION = "" }
      table1p += "<tr><td>"+ table_arr[1][i].item_search_ + "</td><td>"+ table_arr[1][i].item_name_ + "</td><td>" + table_arr[1][i].item_planedquantity + "</td><td>"+ table_arr[1][i].ITEM_ACTUALQUANTITY + "</td><td>"+ table_arr[1][i].ITEM_PLANQUANTITY + "</td><td>"+ table_arr[1][i].ITEM_NOTES_ + "</td><td>" + table_arr[1][i].ITEM_RECTIFICATION + "</td></tr>";
    }
    table1p += "</table>";

    var table2 = "<table border='1'><tr><td>序号</td><td>设备名称</td><td>规格</td><td>本周进退场</td><td>累计数量</td><td>状态描述</td></tr>";
    for( let i=0;i<table_arr[2].length;i++){
      table2 += "<tr><td>" + (i+1) + "</td><td>"+ table_arr[2][i].ITEM_EQUIPMENT + "</td><td>"+ table_arr[2][i].ITEM_规格 + "</td><td>" + table_arr[2][i].ITEM_QUANTITY + "</td><td>"+ table_arr[2][i].ITEM_TOTAL + "</td><td>"+ table_arr[2][i].ITEM_状态 + "</td></tr>";
    }
    table2 += "</table>";

    var table3 = "<table border='1'><tr><td>序号</td><td>材料名称</td><td>单位</td><td>总量</td><td>本周进场</td><td>累计进场</td><td>累计进场比例</td></tr>";    
    for( let i=0;i<table_arr[3].length;i++){
      table3 += "<tr><td>" + (i+1) + "</td><td>"+ table_arr[3][i].ITEM_材料名称 + "</td><td>"+ table_arr[3][i].ITEM_单位 + "</td><td>" + table_arr[3][i].ITEM_总量 + "</td><td>"+ table_arr[3][i].ITEM_本周进场 + "</td><td>"+ table_arr[3][i].ITEM_累计进场 + "</td><td>"+ table_arr[3][i].ITEM_累计进场比例 + "</td></tr>";
    }
    table3 += "</table>";

    var table4 = "<table border='1'><tr><td>序号</td><td>检测内容</td><td>单位</td><td>本周检测</td><td>检测累计情况</td><td>结论</td></tr>";
    for( let i=0;i<table_arr[4].length;i++){
      table4 += "<tr><td>" + (i+1) + "</td><td>"+ table_arr[4][i].ITEM_检测内容 + "</td><td>"+ table_arr[4][i].ITEM_单位 + "</td><td>" + table_arr[4][i].ITEM_本周检测 + "</td><td>"+ table_arr[4][i].ITEM_检测累计情况 + "</td><td>"+ table_arr[4][i].ITEM_结论 + "</td></tr>";
    }
    table4 += "</table>";
    
    var table5 = "<table border='1'><tr><td>序号</td><td>类别</td><td>人数</td><td>备注</td></tr>";
    var tlk5_sum = 0;
    for( let i=0;i<table_arr[5].length;i++){
      table5 += "<tr><td>" + (i+1) + "</td><td>"+ table_arr[5][i].ITEM_WORKER_CLASS + "</td><td>"+ table_arr[5][i].ITEM_QUANTITY + "</td><td>" + table_arr[5][i].ITEM_备注 + "</td></tr>";
      tlk5_sum += table_arr[5][i].ITEM_QUANTITY;
    }
    table5 += "<tr><td colspan='2'>合计</td><td>" + tlk5_sum + "</td><td></td></tr></table>";
    
    const safety = sqlResult[8][0].ITEM_安全;
    const quality = sqlResult[8][0].ITEM_质量;
    const contact = sqlResult[8][0].ITEM_联系单;
    const method = sqlResult[8][0].ITEM_问题落实;
    const consultation1 = sqlResult[8][0].ITEM_影响协商;
    const picture1 = sqlResult[8][0].item_照片;
    const displacement = sqlResult[8][0].ITEM_沉降位移;

    const consultation = consultation1.replace(/src="(.+?)"/g, src => {
      return src.replace(/\//g, '\\').replace(/src="/,'src="file:///C:\\cnqisoft\\bin\\apache-tomcat-7.0.57\\webapps');
    })
    const picture = picture1.replace(/src="(.+?)"/g, src => {
      console.log(src.replace(/\//g, '\\').replace(/src="/,'src="file:///C:\\cnqisoft\\bin\\apache-tomcat-7.0.57\\webapps'));
      return src.replace(/\//g, '\\').replace(/src="/,'src="file:///C:\\cnqisoft\\bin\\apache-tomcat-7.0.57\\webapps');
    })

    //正则替换
    var reg = [
      {
        relus: /_title_/g,
        match: title
      },
      {
        relus: /_week_/g,
        match: week
      },
      {
        relus: /_weekstartYear_/g,
        match: weekstartYear
      },
      {
        relus: /_weekstartMonth_/g,
        match: weekstartMonth
      },
      {
        relus: /_weekstartDay_/g,
        match: weekstartDay
      },
      {
        relus: /_weekfinishYear_/g,
        match: weekfinishYear
      },
      {
        relus: /_weekfinishMonth_/g,
        match: weekfinishMonth
      },
      {
        relus: /_weekfinishDay_/g,
        match: weekfinishDay
      },
      {
        relus: /_deptname_/g,
        match: deptname
      },
      {
        relus: /_date_/g,
        match: moment().format("YYYY年MM月DD日")
      },
      {
        relus: /_table1_/g,
        match: table1
      },{
        relus: /_table1p_/g,
        match: table1p
      },
      {
        relus: /_table2_/g,
        match: table2
      },
      {
        relus: /_table3_/g,
        match: table3
      },
      {
        relus: /_table4_/g,
        match: table4
      },
      {
        relus: /_table5_/g,
        match: table5
      },
      {
        relus: /_safety_/g,
        match: safety
      },
      {
        relus: /_quality_/g,
        match: quality
      },
      {
        relus: /_contact_/g,
        match: contact
      },
      {
        relus: /_method_/g,
        match: method
      },
      {
        relus: /_consultation_/g,
        match: consultation
      },
      {
        relus: /_picture_/g,
        match: picture
      },
      {
        relus: /_displacement_/g,
        match: displacement
      }
    ];
    //传参到html
    create.createPDFProtocolFile(html, options, reg, res);
  }
})
app.listen(3004);