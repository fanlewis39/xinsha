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
        val: '设备名称',
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
        val: '规格',
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
        val: '本周进退场',
        opts: {
            align: 'center',
            color: '000000',
            cellColWidth: 2000,
            sz: font_size_table,
            shd: {
                fill: 'ffffff',
                themeFill: 'text1'
            }
        }
    }, {
        val: '累计数量',
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
        val: '状态描述',
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