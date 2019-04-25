var async = require('async')
var officegen = require('officegen');
var http = require('http');
var url = require('url');
var client = require('mysql');
var table_one = clearCacheAndRequire('./table/one')
var table_one_style = require('./table/one_style')
var imageurlHeader = "../../../../";
var sqlResult, sqlResult1, sqlResult2;
var font_size = 17;
var font_size_table = 30;
String.prototype.replaceAll = function (s1, s2) {
  return this.replace(new RegExp(s1, "gm"), s2)
}

function clearCacheAndRequire(url) {
  delete require.cache[require.resolve(url)]
  return file = require(url)
}

http.createServer(function (request, response) {

  var fs = require('fs')
  var path = require('path')

  var outDir = path.join(__dirname, '../tmp/')

  // var themeXml = fs.readFileSync(path.resolve(__dirname, 'themes/testTheme.xml'), 'utf8')

  var docx = officegen({
    type: 'docx',
    orientation: 'portrait',
    pageMargins: { top: 1000, left: 1000, bottom: 1000, right: 1000 }
    // The theme support is NOT working yet...
    // themeXml: themeXml
  })
  // Remove this comment in case of debugging Officegen:
  // officegen.setVerboseMode ( true )

  docx.on('error', function (err) {
    console.log(err)
  })

  var data = url.parse(request.url, true).query;
  var port = ''
  if (data.url == '192.168.0.15') {
    port = '3307'
  } else {
    port = '3304'
  }

  if (data.url && data.docid && data.projectId) {
    var connection = client.createConnection({
      host: data.url,
      user: "root",
      password: "123456",
      port: port,
      database: 'pm',
      multipleStatements: true
    })
    connection.connect();
    var sql1 = "select domainid,item_week,date_format(item_weekstart,'%Y-%m-%d') as item_weekstart ,date_format(item_weekstart,'%Y') as item_weekstartYear ,date_format(item_weekstart,'%m') as item_weekstartMonth,date_format(item_weekstart,'%d') as item_weekstartDay, date_format(item_weekfinish,'%Y-%m-%d') as item_weekfinish ,date_format(item_weekfinish,'%Y') as item_weekfinishYear ,date_format(item_weekfinish,'%m') as item_weekfinishMonth,date_format(item_weekfinish,'%d') as item_weekfinishDay from tlk_weeklyreport_xs where id='" + data.docid + "'";
    var sql2 = ";select domainid,item_name,item_deptname from tlk_projectsetup where item_PROJECTID = '" + data.projectId + "'";
    var renwujindu_sql = ";select item_search_,item_name_,item_planedquantity,ITEM_ACTUALQUANTITY,ITEM_THEASSIGNMENT,ITEM_PLANQUANTITY,ITEM_NOTES_,ITEM_RECTIFICATION,FORMAT(ITEM_PERCENTCOMPLETE_,2) as ITEM_PERCENTCOMPLETE from tlk_weekreport where parent = '" + data.docid + "' and (ITEM_SUMMARY_ = 1 or ITEM_ACTUALQUANTITY > 0 or (ITEM_THEASSIGNMENT <> '' and ITEM_THEASSIGNMENT is not null )) order by item_search_ asc"
    var shebeizhuangkuang_sql = ";select DOMAINID,ITEM_EQUIPMENT,`ITEM_规格`,FORMAT(ITEM_QUANTITY,0) as ITEM_QUANTITY,ITEM_TOTAL,`ITEM_状态` from tlk_equipment where parent = '" + data.docid + "'"
    var cailiaozhuangkuang_sql = ";select DOMAINID,`ITEM_材料名称`,`ITEM_单位`,`ITEM_总量`,`ITEM_本周进场`,`ITEM_累计进场`,`ITEM_累计进场比例` from `tlk_材料状况` where parent = '" + data.docid + "'"
    var jianceqingkuang_sql = ";select domainid,`ITEM_检测内容`,`ITEM_单位`,`ITEM_本周检测`,`ITEM_检测累计情况`,`ITEM_结论` from `tlk_检测情况` where parent = '" + data.docid + "'";
    var renyuanpeizhi_sql = ";select DOMAINID,ITEM_WORKER_CLASS,ITEM_QUANTITY,`ITEM_备注` from tlk_worker where parent = '" + data.docid + "'";
    var anquanzhiliangneiyye_sql = ";select DOMAINID,`ITEM_安全`,`ITEM_质量`,`ITEM_联系单`,`ITEM_问题落实`,`ITEM_影响协商`,item_照片,`ITEM_沉降位移` from tlk_weeklyreport_xs where id='" + data.docid + "'"
    var yanchikaigong_sql = ";select item_search_,item_name_,item_planedquantity,ITEM_ACTUALQUANTITY,ITEM_THEASSIGNMENT,ITEM_PLANQUANTITY,ITEM_NOTES_,ITEM_RECTIFICATION,FORMAT(ITEM_PERCENTCOMPLETE_,2) as ITEM_PERCENTCOMPLETE from tlk_weekreport where parent = '" + data.docid + "' and ITEM_SUMMARY_ = 0 and ITEM_ACTUALQUANTITY = 0 and (ITEM_THEASSIGNMENT = '' or ITEM_THEASSIGNMENT is null ) order by item_search_ asc"
    var sql = sql1 + sql2 + renwujindu_sql + shebeizhuangkuang_sql + cailiaozhuangkuang_sql + jianceqingkuang_sql + renyuanpeizhi_sql + anquanzhiliangneiyye_sql + yanchikaigong_sql;
    connection.query(sql, (err, result) => {
      if (err) {
        console.log('[SELECT ERROR]-', err.message);
        return;
      }
      sqlResult = [];
      sqlResult = result;
      // console.log('sqlResult',sqlResult[2])
      docx_a();
    })

    connection.end();
  } else {
    makeDocx_error();
  }


  //mysql的连接以及数据获取
  function docx_a() {
    // var table = clearCacheAndRequire('./table/one')
    //第一个表
    var tab1_no = 'XS-ZB-0' + sqlResult[0][0].item_week;
    var tab1_time = sqlResult[0][0].item_weekstartYear + '.' + sqlResult[0][0].item_weekstartMonth + '.' + sqlResult[0][0].item_weekstartDay + '-' + sqlResult[0][0].item_weekfinishYear + '.' + sqlResult[0][0].item_weekfinishMonth + '.' + sqlResult[0][0].item_weekfinishDay;
    var tab1_unit = '中交第四航务工程局有限公司' + sqlResult[1][0].item_name + sqlResult[1][0].item_deptname;
    var table = [
      [{
          val: '编号',
          opts: {
              cellColWidth: 1000,
              color: '000000',
              sz: font_size_table,
              shd: {
                  fill: '000000',
                  themeFill: 'text1'
              }
          }
      }, {
          val: tab1_no,
          opts: {
              align: 'center',
              color: '000000',
              cellColWidth: 2500,
              sz: font_size_table,
              shd: {
                  fill: 'ffffff',
                  themeFill: 'text1'
              }
          }
      }, {
          val: '周期',
          opts: {
              color: '000000',
              cellColWidth: 1000,
              align: 'center',
              sz: font_size_table,
              shd: {
                  fill: '92CDDC',
                  themeFill: 'text1'
              }
          }
      }, {
          val: tab1_time,
          cellColWidth: 1000,
          opts: {
              color: '000000',
              align: 'center',
              sz: font_size_table,
              shd: {
                  fill: '92CDDC',
                  themeFill: 'text1'
              }
          }
      }],
      ['主送单位','广州港工程管理有限公司','发文单位',tab1_unit]
  ]

    var renwujindu = clearCacheAndRequire('./table/renwujindu')
    for (var i in sqlResult[2]) {
      var json = sqlResult[2][i];
      var arr = [json.item_search_, json.item_name_, json.item_planedquantity.toString(), json.ITEM_ACTUALQUANTITY.toString(), real2val(json.ITEM_THEASSIGNMENT), json.ITEM_PLANQUANTITY.toString(), json.ITEM_NOTES_, json.ITEM_RECTIFICATION, json.ITEM_PERCENTCOMPLETE];
      renwujindu.push(arr);
    }

    var shebeizhuangkuang = clearCacheAndRequire('./table/shebeizhuangkuang')
    var num_1 = 1;
    for (var i in sqlResult[3]) {
      var json = sqlResult[3][i];
      var arr = [num_1++, json.ITEM_EQUIPMENT, json.ITEM_规格, json.ITEM_QUANTITY, json.ITEM_TOTAL, json.ITEM_状态];
      shebeizhuangkuang.push(arr);
    }

    var cailiaozhuangkuang = clearCacheAndRequire('./table/cailiaozhuangkuang')
    var num_2 = 1;
    for (var i in sqlResult[4]) {
      var json = sqlResult[4][i];
      var arr = [num_2++, json.ITEM_材料名称, json.ITEM_单位, json.ITEM_总量, json.ITEM_本周进场, json.ITEM_累计进场, json.ITEM_累计进场比例];
      cailiaozhuangkuang.push(arr);
    }

    var jianceqingkuang = clearCacheAndRequire('./table/jianceqingkuang')
    var num_3 = 1;
    for (var i in sqlResult[5]) {
      var json = sqlResult[5][i];
      var arr = [num_3++, json.ITEM_检测内容, json.ITEM_单位, json.ITEM_本周检测, json.ITEM_检测累计情况, json.ITEM_结论];
      jianceqingkuang.push(arr);
    }

    var renyuanpeizhi = clearCacheAndRequire('./table/renyuanpeizhi')
    var num_4 = 1;
    for (var i in sqlResult[6]) {
      var json = sqlResult[6][i];
      var arr = [num_4++, json.ITEM_WORKER_CLASS, json.ITEM_QUANTITY, json.ITEM_备注];
      renyuanpeizhi.push(arr);
    }

    var yanchikaigong = clearCacheAndRequire('./table/yanchikaigong')
    for (var i in sqlResult[8]) {
      var json = sqlResult[8][i];
      var arr = [json.item_search_, json.item_name_, json.item_planedquantity.toString(), json.ITEM_ACTUALQUANTITY.toString(), json.ITEM_PLANQUANTITY.toString(), json.ITEM_NOTES_, json.ITEM_RECTIFICATION];
      yanchikaigong.push(arr);
    }

    var anquan = string2html(sqlResult[7][0].ITEM_安全)
    var zhiliang = string2html(sqlResult[7][0].ITEM_质量)
    var lianxidan = string2html(sqlResult[7][0].ITEM_联系单)
    var wentiluoshi = string2html(sqlResult[7][0].ITEM_问题落实)
    var yingxiangxieshang = string2html(sqlResult[7][0].ITEM_影响协商)
    var photo = string2html(sqlResult[7][0].item_照片)
    var chengjiangweiyi = string2html(sqlResult[7][0].ITEM_沉降位移)

    var pObj = docx.createP()

    sqlResult1 = sqlResult[0];
    sqlResult2 = sqlResult[1];
  
    //header
    var header = docx.getHeader().createP();
    header.addText('    中交第四航务工程局有限公司');
    header.addHorizontalLine()

    //footer
    var footer = docx.getFooter().createP({ align: 'center' });
    footer.addText(sqlResult2[0].item_name)

    //头页
    pObj = docx.createP({ align: 'center' })
    pObj.addText(sqlResult2[0].item_name, { font_size: font_size })
    pObj = docx.createP()
    pObj = docx.createP()
    pObj = docx.createP()
    pObj = docx.createP()
    pObj = docx.createP({ align: 'center' })
    pObj.addText('施工周报', { font_size: 35 })
    pObj = docx.createP()
    pObj = docx.createP({ align: 'center' })
    pObj.addText('（第  ', { font_size: font_size })
    pObj.addText("0" + sqlResult1[0].item_week, { bold: true, underline: true, font_size: font_size })
    pObj.addText('  期，编号：', { font_size: font_size })
    pObj.addText('XS-ZB-0' + sqlResult1[0].item_week, { bold: true, underline: true, font_size: font_size })
    pObj.addText('）', { font_size: font_size })
    pObj = docx.createP({ align: 'center' })
    pObj.addText('（', { font_size: font_size })
    pObj.addText('' + sqlResult1[0].item_weekstartYear, { bold: true, underline: true, font_size: font_size })
    pObj.addText('年', { font_size: font_size })
    pObj.addText('' + sqlResult1[0].item_weekstartMonth, { bold: true, underline: true, font_size: font_size })
    pObj.addText('月', { font_size: font_size })
    pObj.addText('' + sqlResult1[0].item_weekstartDay, { bold: true, underline: true, font_size: font_size })
    pObj.addText('日至', { font_size: font_size })
    pObj.addText('' + sqlResult1[0].item_weekfinishYear, { bold: true, underline: true, font_size: font_size })
    pObj.addText('年', { font_size: font_size })
    pObj.addText('' + sqlResult1[0].item_weekfinishMonth, { bold: true, underline: true, font_size: font_size })
    pObj.addText('月', { font_size: font_size })
    pObj.addText('' + sqlResult1[0].item_weekfinishDay, { bold: true, underline: true, font_size: font_size })
    pObj.addText('日）', { font_size: font_size })
    pObj = docx.createP()
    pObj = docx.createP()
    pObj = docx.createP({ align: 'center' })
    pObj.addText('审批：_______', { font_size: font_size })
    // pObj.addText('     ', { bold: true, underline: true })
    pObj = docx.createP({ align: 'center' })
    pObj.addText('编制：_______', { font_size: font_size })
    // pObj.addText('     ', { bold: true, underline: true })
    pObj = docx.createP()
    pObj = docx.createP()
    pObj = docx.createP()
    pObj = docx.createP({ align: 'center' })
    pObj.addText('中交第四航务工程局有限公司' + sqlResult2[0].item_name, { font_size: font_size })
    pObj.addText(sqlResult2[0].item_deptname, { font_size: font_size })
    pObj = docx.createP({ align: 'center' })
    pObj.addText('编制日期：' + sqlResult1[0].item_weekfinish, { font_size: font_size })
    docx.putPageBreak()
    pObj = docx.createP()

    var tableStyle = table_one_style;

    var data1 = [
      {
        type: "table",
        val: table,
        opt: tableStyle
      }, {
        type: 'createp'
      }, {
        type: "text",
        val: "一、上周施工概要及工程进度情况",
        opt: { bold: true, font_size: font_size }
      }, {
        type: "text",
        val: "1、上周施工概要",
        opt: { font_size: 15 }
      }, {
        type: "table",
        val: renwujindu,
        opt: tableStyle
      }, {
        type: 'createp'
      }, {
        type: "text",
        val: "2、延迟开工任务",
        opt: { font_size: 15 }
      }, {
        type: "table",
        val: yanchikaigong,
        opt: tableStyle
      }, {
        type: 'createp'
      }, {
        type: "text",
        val: "二、设备/材料状况",
        opt: { bold: true, font_size: font_size }
      }, {
        type: "text",
        val: "2、设备状况",
        opt: { font_size: 15 }
      }, {
        type: "table",
        val: shebeizhuangkuang,
        opt: tableStyle
      }, {
        type: 'createp'
      }, {
        type: "text",
        val: "2、材料状况",
        opt: { font_size: 15 }
      }, {
        type: "table",
        val: cailiaozhuangkuang,
        opt: tableStyle
      }, {
        type: 'createp'
      }, {
        type: "text",
        val: "3、检测情况",
        opt: { font_size: font_size }
      }, {
        type: "table",
        val: jianceqingkuang,
        opt: tableStyle
      }, {
        type: 'createp'
      }, {
        type: "text",
        val: "三、人员配置状况",
        opt: { bold: true, font_size: font_size }
      }, {
        type: "table",
        val: renyuanpeizhi,
        opt: tableStyle
      }, {
        type: 'createp'
      }, {
        type: "text",
        val: "四、安全、质量、内业情况",
        opt: { bold: true, font_size: font_size }
      }
    ]
    docx.createByJson(data1);


    var json_test = {
      "safe": anquan,
      "zhiliang": zhiliang,
      "lianxidan": lianxidan,
      "problem": wentiluoshi,
      'yingxiang': yingxiangxieshang,
      "photo": photo,
      'down': chengjiangweiyi
    }
    addimage2(json_test);

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
    docx.generate(response);
  }

  var fileName = 'weekreport';
  response.writeHead(200, {
    'Access-Control-Allow-Origin': '*',
    'Content-Type': 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
    'Content-disposition': 'attachment; filename=' + fileName + '.docx',
    'Access-Control-Allow-Credentials': 'true',
    'Access-Control-Allow-Methods': 'PUT,POST,GET,DELETE,OPTIONS'
  });

  function makeDocx_error() {
    pObj = docx.createP({ align: 'center' })
    pObj.addText('请进入表单后下载', { font_size: 35 })
    pObj = docx.createP()
    var out = fs.createWriteStream('example.docx')
    out.on('error', function (err) {
      console.log(err)
    })
    async.parallel([
      function (done) {
        out.on('close', function () {
          console.log('Finish to create a DOCX_error file.')
          done(null)
        })
        docx.generate(out)
      }
    ], function (err) {
      if (err) {
        console.log('error: ' + err)
      } // Endif.
    })
    docx.generate(response);
  }

  function string2html(string) {
    var arr = [];
    var imgObjArr2 = html2src(string);
    for (var i in imgObjArr2) {
      arr.push(imgObjArr2[i])
    }
    return arr;
  }

  function html2text(string) {
    if (string) {
      return string = JSON.parse(JSON.stringify(string).replace(/<\/?.+?\/?>/g, ""));
    }
  }
  function html2src(string) {
    var ObjArr = [];
    if (string) {
      var imgReg = /<img.*?(?:>|\/>)/gi;
      var pReg = /<p.*?>.*?<\/p>/g;
      var srcReg = /src=[\'\"]?([^\'\"]*)[\'\"]?/i;
      var altReg = /alt=[\'\"]?([^\'\"]*)[\'\"]?/i;
      var arr = string.match(imgReg);
      var arr_p = string.match(pReg);
      function imgObj(src, alt) {
        var obj = {
          src: src,
          alt: alt,
        }
        return obj
      }
      function pObj(val) {
        var obj = {
          val: val
        }
        return obj;
      }
      if (arr_p != null) {
        for (var i in arr_p) {
          var arr = arr_p[i].match(imgReg);
          ObjArr.push(new pObj(html2text(arr_p[i])))
          if (arr != null) {
            for (var i = 0; i < arr.length; i++) {
              var src = arr[i].match(srcReg);
              var alt = arr[i].match(altReg);
              //获取alt内容
              if (alt[1]) {
                var obj = new imgObj(src[1], alt[1])
                ObjArr.push(obj);
              } else {
                var obj = new imgObj(scr[1], '')
                ObjArr.push(obj);
              }
            }
          }
        }
      }
    }
    return ObjArr;
  }

  function addimage2(json) {
    var path = require('path');
    var imgTitle = "";
    pObj = docx.createP();
    var a = ''
    for (var i in json) {
      if (i == 'problem') {
        pObj = docx.createP({ align: 'left' });
        pObj.addText('五、上周问题落实情况', { bold: true, font_size: font_size });
        a = i;
      } else if (i == 'safe') {
        pObj = docx.createP({ align: 'left' });
        pObj.addText('1、安全生产、文明施工情况', { bold: true, font_size: font_size });
        a = i;
      } else if (i == 'zhiliang') {
        pObj = docx.createP({ align: 'left' });
        pObj.addText('2、施工质量情况', { bold: true, font_size: font_size });
        a = i;
      } else if (i == 'lianxidan') {
        pObj = docx.createP({ align: 'left' });
        pObj.addText('3、主要联系单', { bold: true, font_size: font_size });
        a = i;
      } else if (i == 'photo') {
        pObj = docx.createP({ align: 'left' });
        pObj.addText('八、照片', { bold: true, font_size: font_size });
        pObj = docx.createP();
        a = i;
      } else if (i == 'yingxiang') {
        pObj = docx.createP({ align: 'left' });
        pObj.addText('七、主要影响因素及需协商解决问题', { bold: true, font_size: font_size });
        pObj = docx.createP();
        a = i;
      } else {
        docx.putPageBreak()
        pObj = docx.createP();
        pObj.addText('九、沉降位移观测', { bold: true, font_size: font_size });
        pObj = docx.createP();
        a = i;
      }

      if (!json[i].length) {
        pObj.addText('无', { font_size: font_size, align: 'center' });
      } else {
        var jsonCopy = json[i];
        for (var i in jsonCopy) {
          if (jsonCopy[i].val) {
            pObj = docx.createP();
            pObj.addText(jsonCopy[i].val.replaceAll('&nbsp;', ''), { font_size: font_size });
          }
          if (jsonCopy[i].src) {
            pObj = docx.createP({ align: 'center' });
            var src = jsonCopy[i].src;

            if (a == 'down') {
              pObj.addImage(path.resolve(__dirname, imageurlHeader + src), { cx: 600, cy: 730, align: 'center' });
            } else {
              pObj.addImage(path.resolve(__dirname, imageurlHeader + src), { cx: 400, cy: 266, align: 'center' });
            }
          }
        }
      }
      pObj = docx.createP();
    }
  }
  //真实值转显示值
  function real2val(real) {
    if (real == 1) {
      return '进场准备'
    } else if (real == 2) {
      return '计划性停工'
    } else if (real == 3) {
      return '意外性停工'
    }
  }
}).listen(4004);

