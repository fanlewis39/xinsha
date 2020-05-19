let font_size_table = 20;
let colWidth = 2000;
let colWidth2 = 1500;
let jianceqingkuang = [
    [{
        val: '序号',
        opts: {
            cellColWidth: colWidth2,
            color: '000000',
            sz: font_size_table,
            shd: {
                fill: '000000',
                themeFill: 'text1'
            }
            // fontFamily: 'Avenir Book'
        }
    }, {
        val: '检测内容',
        opts: {
            align: 'center',
            color: '000000',
            cellColWidth: colWidth,
            sz: font_size_table,
            shd: {
                fill: 'ffffff',
                themeFill: 'text1'
            }
        }
    }, {
        val: '单位',
        opts: {
            color: '000000',
            cellColWidth: colWidth2,
            align: 'center',
            sz: font_size_table,
            shd: {
                fill: '92CDDC',
                themeFill: 'text1'
                // 'themeFillTint': '80'
            }
        }
    }, {
        val: '本周检测',
        opts: {
            align: 'center',
            color: '000000',
            cellColWidth: colWidth,
            sz: font_size_table,
            shd: {
                fill: 'ffffff',
                themeFill: 'text1'
            }
        }
    }, {
        val: '检测累计情况',
        opts: {
            align: 'center',
            color: '000000',
            cellColWidth: colWidth,
            sz: font_size_table,
            shd: {
                fill: 'ffffff',
                themeFill: 'text1'
            }
        }
    }, {
        val: '结论',
        opts: {
            align: 'center',
            color: '000000',
            cellColWidth: colWidth2,
            sz: font_size_table,
            shd: {
                fill: 'ffffff',
                themeFill: 'text1'
            }
        }
    }]
]

module.exports = jianceqingkuang