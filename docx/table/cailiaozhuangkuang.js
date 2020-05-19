let font_size_table = 20;
let colWidth = 1500;

let shebeizhuangkuang = [
    [{
        val: '序号',
        opts: {
            cellColWidth: colWidth,
            color: '000000',
            sz: font_size_table,
            shd: {
                fill: '000000',
                themeFill: 'text1'
            }
            // fontFamily: 'Avenir Book'
        }
    }, {
        val: '材料名称',
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
        val: '单位',
        opts: {
            color: '000000',
            cellColWidth: 1000,
            align: 'center',
            sz: font_size_table,
            shd: {
                fill: '92CDDC',
                themeFill: 'text1'
                // 'themeFillTint': '80'
            }
        }
    }, {
        val: '总量',
        opts: {
            align: 'center',
            color: '000000',
            cellColWidth: 1000,
            sz: font_size_table,
            shd: {
                fill: 'ffffff',
                themeFill: 'text1'
            }
        }
    }, {
        val: '本周进场',
        opts: {
            align: 'center',
            color: '000000',
            cellColWidth: 1000,
            sz: font_size_table,
            shd: {
                fill: 'ffffff',
                themeFill: 'text1'
            }
        }
    }, {
        val: '累计进场',
        opts: {
            align: 'center',
            color: '000000',
            cellColWidth: 1000,
            sz: font_size_table,
            shd: {
                fill: 'ffffff',
                themeFill: 'text1'
            }
        }
    }, {
        val: '累计进场比例（%）',
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
    }]
]

module.exports = shebeizhuangkuang