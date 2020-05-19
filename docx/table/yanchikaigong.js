let font_size_table = 20;
let colWidth = 1000;

let yanchikaigong = [
    [{
        val: 'WBS',
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
        val: '名称',
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
        val: '本周计划',
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
        val: '实际完成',
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
    },  {
        val: '下周计划',
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
        val: '偏离说明',
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
        val: '纠偏措施',
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
module.exports = yanchikaigong