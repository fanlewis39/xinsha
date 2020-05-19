const officegen = require("officegen");
const http = require("http");
const url = require("url");
const mysql = require("mysql");
const table_one_style = require("./table/one_style");

String.prototype.replaceAll = function(s1, s2) {
    return this.replace(new RegExp(s1, "gm"), s2);
};

function clearCacheAndRequire(url) {
    delete require.cache[require.resolve(url)];
    return (file = require(url));
}

function excludeDuplicates(arr1,arr2) {
    for(let i in arr1){
      for(let j in arr2){
        if(arr1[i]["item_search_"] == (arr2[j]["item_search_"])){
            arr1.splice(i,1);
        }       
        continue;
      }
    }
    return arr1;
  }

http.createServer(function(request, response) {
    let imageurlHeader = "../../../../..";
    let sqlResult, sqlResult1, sqlResult2;
    let font_size = 17;
    let font_size_table = 30;
    let docx = officegen({
        type: "docx",
        orientation: "portrait",
        pageMargins: { top: 1000, left: 1000, bottom: 1000, right: 1000 }
        // The theme support is NOT working yet...
        // themeXml: themeXml
    });
    // Remove this comment in case of debugging Officegen:
    // officegen.setVerboseMode ( true )

    docx.on("error", function(err) {
        console.log(err);
    });

    let data = url.parse(request.url, true).query;

    if (data.docid && data.projectId) {
        let connection = mysql.createConnection({
            host: "localhost",
            user: "root",
            password: "123456",
            port: "3307",
            database: "pm",
            multipleStatements: true
        });
        connection.connect();
        const sql1 = "select domainid,item_week,date_format(item_weekstart,'%Y-%m-%d') as item_weekstart ,date_format(item_weekstart,'%Y') as item_weekstartYear ,date_format(item_weekstart,'%m') as item_weekstartMonth,date_format(item_weekstart,'%d') as item_weekstartDay, date_format(item_weekfinish,'%Y-%m-%d') as item_weekfinish ,date_format(item_weekfinish,'%Y') as item_weekfinishYear ,date_format(item_weekfinish,'%m') as item_weekfinishMonth,date_format(item_weekfinish,'%d') as item_weekfinishDay from tlk_weeklyreport_xs where id='" + data.docid + "'";
        const sql2 = ";select domainid,item_name,item_deptname from tlk_projectsetup where item_PROJECTID = '" + data.projectId + "'";
        //正常进行任务
        const zhengchangjinxing_sql = ";select item_search_,item_name_,item_总工程量,item_计量单位,item_planedquantity,ITEM_ACTUALQUANTITY,ITEM_PLANQUANTITY,ITEM_计划累计工程量 as ITEM_计划累计完成工程量,ITEM_计划累计完成 as ITEM_计划累计完成百分比,ITEM_实际累计完成 as ITEM_实际累计完成工程量,FORMAT(ITEM_PERCENTCOMPLETE_,2) as ITEM_实际累计完成百分比 FROM tlk_weekreport where parent = '" + data.docid + "' and item_percentcomplete_ > 0 and item_percentcomplete_ < 100 and item_percentcomplete_ >= (item_计划累计完成*100 - 1) and item_percentcomplete_ <= (item_计划累计完成*100 + 1)";
        //进度滞后任务
        const jinduzhihou_sql = ";select item_search_,item_name_,item_总工程量,item_计量单位,item_planedquantity,ITEM_ACTUALQUANTITY,ITEM_PLANQUANTITY,ITEM_计划累计工程量 as ITEM_计划累计完成工程量,ITEM_计划累计完成 as ITEM_计划累计完成百分比,ITEM_实际累计完成 as ITEM_实际累计完成工程量,FORMAT(ITEM_PERCENTCOMPLETE_,2) as ITEM_实际累计完成百分比 FROM tlk_weekreport where parent = '" + data.docid + "' and item_percentcomplete_ > 0 and item_percentcomplete_ < 100 and item_percentcomplete_ < (item_计划累计完成*100 - 1)";
        //进度超前任务
        const jinduchaoqian_sql = ";select item_search_,item_name_,item_总工程量,item_计量单位,item_planedquantity,ITEM_ACTUALQUANTITY,ITEM_PLANQUANTITY,ITEM_计划累计工程量 as ITEM_计划累计完成工程量,ITEM_计划累计完成 as ITEM_计划累计完成百分比,ITEM_实际累计完成 as ITEM_实际累计完成工程量,FORMAT(ITEM_PERCENTCOMPLETE_,2) as ITEM_实际累计完成百分比 FROM tlk_weekreport where parent = '" + data.docid + "' and item_percentcomplete_ > 0 and item_percentcomplete_ < 100 and item_percentcomplete_ > (item_计划累计完成*100 + 1)";
        //延迟开工初步筛选
        const yanchikaigong_sql = ";select item_search_,item_name_,item_planedquantity,ITEM_ACTUALQUANTITY,ITEM_THEASSIGNMENT,ITEM_PLANQUANTITY,ITEM_NOTES_,ITEM_RECTIFICATION,item_计量单位 FROM tlk_weekreport where parent = '" + data.docid + "' and ITEM_ACTUALQUANTITY = 0 and (ITEM_THEASSIGNMENT = '' or ITEM_THEASSIGNMENT is null )";
        const shebeizhuangkuang_sql = ";select DOMAINID,ITEM_EQUIPMENT,`ITEM_规格`,FORMAT(ITEM_QUANTITY,0) as ITEM_QUANTITY,ITEM_TOTAL,`ITEM_状态` from tlk_equipment where parent = '" + data.docid + "'";
        const cailiaozhuangkuang_sql = ";select DOMAINID,`ITEM_材料名称`,`ITEM_单位`,`ITEM_总量`,`ITEM_本周进场`,`ITEM_累计进场`,`ITEM_累计进场比例` from `tlk_材料状况` where parent = '" + data.docid + "'";
        const jianceqingkuang_sql = ";select domainid,`ITEM_检测内容`,`ITEM_单位`,`ITEM_本周检测`,`ITEM_检测累计情况`,`ITEM_结论` from `tlk_检测情况` where parent = '" + data.docid + "'";
        const renyuanpeizhi_sql = ";select DOMAINID,ITEM_WORKER_CLASS,ITEM_QUANTITY,`ITEM_备注` from tlk_worker where parent = '" + data.docid + "'";
        const anquanzhiliangneiyye_sql = ";select DOMAINID,`ITEM_安全`,`ITEM_质量`,`ITEM_联系单`,`ITEM_问题落实`,`ITEM_影响协商`,item_照片,`ITEM_沉降位移` from tlk_weeklyreport_xs where id='" + data.docid + "'";
        const sql = sql1 + sql2 + zhengchangjinxing_sql + jinduzhihou_sql + jinduchaoqian_sql + yanchikaigong_sql + shebeizhuangkuang_sql + cailiaozhuangkuang_sql + jianceqingkuang_sql + renyuanpeizhi_sql + anquanzhiliangneiyye_sql;
        connection.query(sql, (err, result) => {
            if (err) {
                console.log("[SELECT ERROR]-", err.message);
                return;
            }
            sqlResult = [];
            sqlResult = result;
            docx_a();
        });

        connection.end();
    } else {
        makeDocx_error();
    }

    //mysql的连接以及数据获取
    function docx_a() {
        //第一个表
        let tab1_no = "XS-ZB-0" + sqlResult[0][0].item_week;
        let tab1_time = sqlResult[0][0].item_weekstartYear + "." + sqlResult[0][0].item_weekstartMonth + "." + sqlResult[0][0].item_weekstartDay + "-" + sqlResult[0][0].item_weekfinishYear + "." + sqlResult[0][0].item_weekfinishMonth + "." + sqlResult[0][0].item_weekfinishDay;
        let tab1_unit = "中交第四航务工程局有限公司" + sqlResult[1][0].item_name + sqlResult[1][0].item_deptname;
        let table = [
            [{
                    val: "编号",
                    opts: {
                        cellColWidth: 1000,
                        color: "000000",
                        sz: font_size_table,
                        shd: {
                            fill: "000000",
                            themeFill: "text1"
                        }
                    }
                },
                {
                    val: tab1_no,
                    opts: {
                        align: "center",
                        color: "000000",
                        cellColWidth: 2500,
                        sz: font_size_table,
                        shd: {
                            fill: "ffffff",
                            themeFill: "text1"
                        }
                    }
                },
                {
                    val: "周期",
                    opts: {
                        color: "000000",
                        cellColWidth: 1000,
                        align: "center",
                        sz: font_size_table,
                        shd: {
                            fill: "92CDDC",
                            themeFill: "text1"
                        }
                    }
                },
                {
                    val: tab1_time,
                    cellColWidth: 1000,
                    opts: {
                        color: "000000",
                        align: "center",
                        sz: font_size_table,
                        shd: {
                            fill: "92CDDC",
                            themeFill: "text1"
                        }
                    }
                }
            ],
            ["主送单位", "广州港工程管理有限公司", "发文单位", tab1_unit]
        ];

        let zhengchangjinxing = clearCacheAndRequire("./table/renwujindu");
        for (let i in sqlResult[2]) {
            let json = sqlResult[2][i];
            let arr = [
                json.item_search_,
                json.item_name_,
                json.item_总工程量 + json.item_计量单位,
                json.item_planedquantity + json.item_计量单位,
                json.ITEM_ACTUALQUANTITY + json.item_计量单位,
                json.ITEM_PLANQUANTITY.toString() + json.item_计量单位,
                json.ITEM_计划累计完成工程量 + json.item_计量单位,
                json.ITEM_实际累计完成工程量 + json.item_计量单位,
                (json.ITEM_计划累计完成百分比*100).toFixed(2) + '%',
                json.ITEM_实际累计完成百分比 + '%'
            ];
            zhengchangjinxing.push(arr);
        }

        let jinduzhihou = clearCacheAndRequire("./table/renwujindu");
        for (let i in sqlResult[3]) {
            let json = sqlResult[3][i];
            let arr = [
                json.item_search_,
                json.item_name_,
                json.item_总工程量 + json.item_计量单位,
                json.item_planedquantity + json.item_计量单位,
                json.ITEM_ACTUALQUANTITY + json.item_计量单位,
                json.ITEM_PLANQUANTITY.toString() + json.item_计量单位,
                json.ITEM_计划累计完成工程量 + json.item_计量单位,
                json.ITEM_实际累计完成工程量 + json.item_计量单位,
                (json.ITEM_计划累计完成百分比*100).toFixed(2) + '%',
                json.ITEM_实际累计完成百分比 + '%'
            ];
            jinduzhihou.push(arr);
        }

        let jinduchaoqian = clearCacheAndRequire("./table/renwujindu");
        for (let i in sqlResult[4]) {
            let json = sqlResult[4][i];
            let arr = [
                json.item_search_,
                json.item_name_,
                json.item_总工程量 + json.item_计量单位,
                json.item_planedquantity + json.item_计量单位,
                json.ITEM_ACTUALQUANTITY + json.item_计量单位,
                json.ITEM_PLANQUANTITY.toString() + json.item_计量单位,
                json.ITEM_计划累计完成工程量 + json.item_计量单位,
                json.ITEM_实际累计完成工程量 + json.item_计量单位,
                (json.ITEM_计划累计完成百分比*100).toFixed(2) + '%',
                json.ITEM_实际累计完成百分比 + '%'
            ];
            jinduchaoqian.push(arr);
        }

        excludeDuplicates(sqlResult[5],sqlResult[3])
        let yanchikaigong = clearCacheAndRequire("./table/yanchikaigong");
        for (let i in sqlResult[5]) {
            let json = sqlResult[5][i];
            let arr = [
                json.item_search_,
                json.item_name_,
                json.item_planedquantity.toString(),
                json.ITEM_ACTUALQUANTITY.toString(),
                json.ITEM_PLANQUANTITY.toString(),
                json.ITEM_NOTES_,
                json.ITEM_RECTIFICATION
            ];
            yanchikaigong.push(arr);
        }

        let shebeizhuangkuang = clearCacheAndRequire("./table/shebeizhuangkuang");
        let num_1 = 1;
        for (let i in sqlResult[6]) {
            let json = sqlResult[6][i];
            let arr = [
                num_1++,
                json.ITEM_EQUIPMENT,
                json.ITEM_规格,
                json.ITEM_QUANTITY,
                json.ITEM_TOTAL,
                json.ITEM_状态
            ];
            shebeizhuangkuang.push(arr);
        }

        let cailiaozhuangkuang = clearCacheAndRequire("./table/cailiaozhuangkuang");
        let num_2 = 1;
        for (let i in sqlResult[7]) {
            let json = sqlResult[7][i];
            let arr = [
                num_2++,
                json.ITEM_材料名称,
                json.ITEM_单位,
                json.ITEM_总量,
                json.ITEM_本周进场,
                json.ITEM_累计进场,
                json.ITEM_累计进场比例
            ];
            cailiaozhuangkuang.push(arr);
        }

        let jianceqingkuang = clearCacheAndRequire("./table/jianceqingkuang");
        let num_3 = 1;
        for (let i in sqlResult[8]) {
            let json = sqlResult[8][i];
            let arr = [
                num_3++,
                json.ITEM_检测内容,
                json.ITEM_单位,
                json.ITEM_本周检测,
                json.ITEM_检测累计情况,
                json.ITEM_结论
            ];
            jianceqingkuang.push(arr);
        }

        let renyuanpeizhi = clearCacheAndRequire("./table/renyuanpeizhi");
        let num_4 = 1;
        for (let i in sqlResult[9]) {
            let json = sqlResult[9][i];
            let arr = [
                num_4++,
                json.ITEM_WORKER_CLASS,
                json.ITEM_QUANTITY,
                json.ITEM_备注
            ];
            renyuanpeizhi.push(arr);
        }

        let anquan = string2html(sqlResult[10][0].ITEM_安全);
        let zhiliang = string2html(sqlResult[10][0].ITEM_质量);
        let lianxidan = string2html(sqlResult[10][0].ITEM_联系单);
        let wentiluoshi = string2html(sqlResult[10][0].ITEM_问题落实);
        let yingxiangxieshang = string2html(sqlResult[10][0].ITEM_影响协商);
        let photo = string2html(sqlResult[10][0].item_照片);
        let chengjiangweiyi = string2html(sqlResult[10][0].ITEM_沉降位移);

        let pObj = docx.createP();

        sqlResult1 = sqlResult[0];
        sqlResult2 = sqlResult[1];

        //header
        let header = docx.getHeader().createP();
        header.addText("    中交第四航务工程局有限公司");
        header.addHorizontalLine();

        //footer
        let footer = docx.getFooter().createP({ align: "center" });
        footer.addText(
            sqlResult2[0].item_name + "                         施工周报"
        );

        let fileName = "weekreport0" + sqlResult1[0].item_week;

        //头页
        pObj = docx.createP({ align: "center" });
        pObj.addText(sqlResult2[0].item_name, { font_size: font_size });
        pObj = docx.createP();
        pObj = docx.createP();
        pObj = docx.createP();
        pObj = docx.createP();
        pObj = docx.createP({ align: "center" });
        pObj.addText("施工周报", { font_size: 35 });
        pObj = docx.createP();
        pObj = docx.createP({ align: "center" });
        pObj.addText("（第  ", { font_size: font_size });
        pObj.addText("0" + sqlResult1[0].item_week, {
            bold: true,
            underline: true,
            font_size: font_size
        });
        pObj.addText("  期，编号：", { font_size: font_size });
        pObj.addText("XS-ZB-0" + sqlResult1[0].item_week, {
            bold: true,
            underline: true,
            font_size: font_size
        });
        pObj.addText("）", { font_size: font_size });
        pObj = docx.createP({ align: "center" });
        pObj.addText("（", { font_size: font_size });
        pObj.addText("" + sqlResult1[0].item_weekstartYear, {
            bold: true,
            underline: true,
            font_size: font_size
        });
        pObj.addText("年", { font_size: font_size });
        pObj.addText("" + sqlResult1[0].item_weekstartMonth, {
            bold: true,
            underline: true,
            font_size: font_size
        });
        pObj.addText("月", { font_size: font_size });
        pObj.addText("" + sqlResult1[0].item_weekstartDay, {
            bold: true,
            underline: true,
            font_size: font_size
        });
        pObj.addText("日至", { font_size: font_size });
        pObj.addText("" + sqlResult1[0].item_weekfinishYear, {
            bold: true,
            underline: true,
            font_size: font_size
        });
        pObj.addText("年", { font_size: font_size });
        pObj.addText("" + sqlResult1[0].item_weekfinishMonth, {
            bold: true,
            underline: true,
            font_size: font_size
        });
        pObj.addText("月", { font_size: font_size });
        pObj.addText("" + sqlResult1[0].item_weekfinishDay, {
            bold: true,
            underline: true,
            font_size: font_size
        });
        pObj.addText("日）", { font_size: font_size });
        pObj = docx.createP();
        pObj = docx.createP();
        pObj = docx.createP({ align: "center" });
        pObj.addText("审批：_______", { font_size: font_size });
        pObj = docx.createP({ align: "center" });
        pObj.addText("编制：_______", { font_size: font_size });
        pObj = docx.createP();
        pObj = docx.createP();
        pObj = docx.createP();
        pObj = docx.createP({ align: "center" });
        pObj.addText("中交第四航务工程局有限公司" + sqlResult2[0].item_name, {
            font_size: font_size
        });
        pObj.addText(sqlResult2[0].item_deptname, { font_size: font_size });
        pObj = docx.createP({ align: "center" });
        pObj.addText("编制日期：" + sqlResult1[0].item_weekfinish, {
            font_size: font_size
        });
        docx.putPageBreak();
        pObj = docx.createP();

        let tableStyle = table_one_style;

        let data1 = [{
                type: "table",
                val: table,
                opt: tableStyle
            },
            {
                type: "createp"
            },
            {
                type: "text",
                val: "一、上周施工概要及工程进度情况",
                opt: { bold: true, font_size: font_size }
            },
            {
                type: "text",
                val: "1、正常进行任务",
                opt: { font_size: 15 }
            },
            {
                type: "table",
                val: zhengchangjinxing,
                opt: tableStyle
            },
            {
                type: "createp"
            },
            {
                type: "text",
                val: "2、进度滞后任务",
                opt: { font_size: 15 }
            },
            {
                type: "table",
                val: jinduzhihou,
                opt: tableStyle
            },
            {
                type: "createp"
            },
            {
                type: "text",
                val: "3、进度超前任务",
                opt: { font_size: 15 }
            },
            {
                type: "table",
                val: jinduchaoqian,
                opt: tableStyle
            },
            {
                type: "createp"
            },
            {
                type: "text",
                val: "4、延迟开工任务",
                opt: { font_size: 15 }
            },
            {
                type: "table",
                val: yanchikaigong,
                opt: tableStyle
            },
            {
                type: "createp"
            },
            {
                type: "text",
                val: "二、设备/材料状况",
                opt: { bold: true, font_size: font_size }
            },
            {
                type: "text",
                val: "1、设备状况",
                opt: { font_size: 15 }
            },
            {
                type: "table",
                val: shebeizhuangkuang,
                opt: tableStyle
            },
            {
                type: "createp"
            },
            {
                type: "text",
                val: "2、材料状况",
                opt: { font_size: 15 }
            },
            {
                type: "table",
                val: cailiaozhuangkuang,
                opt: tableStyle
            },
            {
                type: "createp"
            },
            {
                type: "text",
                val: "3、检测情况",
                opt: { font_size: font_size }
            },
            {
                type: "table",
                val: jianceqingkuang,
                opt: tableStyle
            },
            {
                type: "createp"
            },
            {
                type: "text",
                val: "三、人员配置状况",
                opt: { bold: true, font_size: font_size }
            },
            {
                type: "table",
                val: renyuanpeizhi,
                opt: tableStyle
            },
            {
                type: "createp"
            },
            {
                type: "text",
                val: "四、安全、质量、内业情况",
                opt: { bold: true, font_size: font_size }
            }
        ];
        
        for(let i=2;i<6;i++){
            if(sqlResult[i].length == 0){
                data1[3*i-2] = {
                    type: "text",
                    val: "无",
                    opt: { font_size: font_size }
                }
            }
        }
        if(sqlResult[9].length == 0){
            data1[26] = {
                type: "text",
                val: "无",
                opt: { font_size: font_size }
            }
        }
        docx.createByJson(data1);

        let json_test = {
            safe: anquan,
            zhiliang: zhiliang,
            lianxidan: lianxidan,
            problem: wentiluoshi,
            yingxiang: yingxiangxieshang,
            photo: photo,
            down: chengjiangweiyi
        };
        addimage2(json_test);

        docx.generate(response);

        response.writeHead(200, {
            "Content-Type": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
            "Content-disposition": "attachment; filename=" + fileName + ".docx"
        });

        function string2html(string) {
            let arr = [];
            let imgObjArr2 = html2src(string);
            for (let i in imgObjArr2) {
                arr.push(imgObjArr2[i]);
            }
            return arr;
        }

        function html2text(string) {
            if (string) {
                return (string = JSON.parse(
                    JSON.stringify(string).replace(/<\/?.+?\/?>/g, "")
                ));
            }
        }

        function html2src(string) {
            let ObjArr = [];
            if (string) {
                let imgReg = /<img.*?(?:>|\/>)/gi;
                let pReg = /<p.*?>.*?<\/p>/g;
                let srcReg = /src=[\'\"]?([^\'\"]*)[\'\"]?/i;
                let altReg = /alt=[\'\"]?([^\'\"]*)[\'\"]?/i;
                let arr = string.match(imgReg);
                let arr_p = string.match(pReg);

                function imgObj(src, alt) {
                    let obj = {
                        src: src,
                        alt: alt
                    };
                    return obj;
                }

                function pObj(val) {
                    let obj = {
                        val: val
                    };
                    return obj;
                }
                if (arr_p != null) {
                    for (let i in arr_p) {
                        let arr = arr_p[i].match(imgReg);
                        ObjArr.push(new pObj(html2text(arr_p[i])));
                        if (arr != null) {
                            for (let i = 0; i < arr.length; i++) {
                                let src = arr[i].match(srcReg);
                                let alt = arr[i].match(altReg);
                                //获取alt内容
                                if (alt[1]) {
                                    let obj = new imgObj(src[1], alt[1]);
                                    ObjArr.push(obj);
                                } else {
                                    let obj = new imgObj(scr[1], "");
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
            let path = require("path");
            pObj = docx.createP();
            let a = "";
            for (let i in json) {
                if (i == "problem") {
                    pObj = docx.createP({ align: "left" });
                    pObj.addText("五、上周问题落实情况", {
                        bold: true,
                        font_size: font_size
                    });
                    a = i;
                } else if (i == "safe") {
                    pObj = docx.createP({ align: "left" });
                    pObj.addText("1、安全生产、文明施工情况", {
                        bold: true,
                        font_size: font_size
                    });
                    a = i;
                } else if (i == "zhiliang") {
                    pObj = docx.createP({ align: "left" });
                    pObj.addText("2、施工质量情况", {
                        bold: true,
                        font_size: font_size
                    });
                    a = i;
                } else if (i == "lianxidan") {
                    pObj = docx.createP({ align: "left" });
                    pObj.addText("3、主要联系单", { bold: true, font_size: font_size });
                    a = i;
                } else if (i == "photo") {
                    pObj = docx.createP({ align: "left" });
                    pObj.addText("八、照片", { bold: true, font_size: font_size });
                    pObj = docx.createP();
                    a = i;
                } else if (i == "yingxiang") {
                    pObj = docx.createP({ align: "left" });
                    pObj.addText("七、主要影响因素及需协商解决问题", {
                        bold: true,
                        font_size: font_size
                    });
                    pObj = docx.createP();
                    a = i;
                } else {
                    docx.putPageBreak();
                    pObj = docx.createP();
                    pObj.addText("九、沉降位移观测", {
                        bold: true,
                        font_size: font_size
                    });
                    pObj = docx.createP();
                    a = i;
                }

                if (!json[i].length) {
                    pObj.addText("无", { font_size: font_size, align: "center" });
                } else {
                    let jsonCopy = json[i];
                    for (let i in jsonCopy) {
                        if (jsonCopy[i].val) {
                            pObj = docx.createP();
                            pObj.addText(jsonCopy[i].val.replaceAll("&nbsp;", ""), {
                                font_size: font_size
                            });
                        }
                        if (jsonCopy[i].src) {
                            pObj = docx.createP({ align: "center" });
                            let src = jsonCopy[i].src;

                            if (a == "down") {
                                pObj.addImage(path.resolve(__dirname, imageurlHeader + src), {
                                    cx: 600,
                                    cy: 730,
                                    align: "center"
                                });
                            } else {
                                pObj.addImage(path.resolve(__dirname, imageurlHeader + src), {
                                    cx: 400,
                                    cy: 266,
                                    align: "center"
                                });
                            }
                        }
                    }
                }
                pObj = docx.createP();
            }
        }
    }
}).listen(4004);