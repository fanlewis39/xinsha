var pdf = require("html-pdf");
exports.createPDFProtocolFile = function(template, options, reg ,res) {
  /**
  template: html 模板
  options: 配置
  reg: 正则匹配规则
*/
  // 将所有匹配规则在html模板中匹配一遍
  if (reg && Array.isArray(reg)) {
    reg.forEach(item => {
      template = template.replace(item.relus, item.match);
    });
  }
  pdf.create(template, options).toStream((err, stream) => {
    if (err) return res.end(err.stack); 
    res.setHeader("Content-type", "application/pdf");
    stream.pipe(res);
  });;
};
