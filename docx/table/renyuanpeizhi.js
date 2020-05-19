let font_size_table = 20;
let colWidth = 3000;

let shebeizhuangkuang = [
    [{
        val: '序号',
        opts: {
            cellColWidth: 1500,
            color: '000000',
            sz: font_size_table,
            shd: {
                fill: '000000',
                themeFill: 'text1'
            }
            // fontFamily: 'Avenir Book'
        }
    }, {
        val: '类别',
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
        val: '人数',
        opts: {
            color: '000000',
            cellColWidth: colWidth,
            align: 'center',
            sz: font_size_table,
            shd: {
                fill: '92CDDC',
                themeFill: 'text1'
                // 'themeFillTint': '80'
            }
        }
    }, {
        val: '备注',
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
    }]
]

module.exports = shebeizhuangkuang