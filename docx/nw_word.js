const Service = require('node-windows').Service;

let svc = new Service({
    name:"node_word",//服务器名称
    description:'周报文档下载',
    script:'./make_docx.js',
    wait:'1',
    grow:'0.25',
    maxRestarts:'40',
});

//监听安装事件
svc.on('install', () => {
    svc.start();
    console.log('install complete.');
});

//监听卸载事件
svc.on('uninstall', () => {
    console.log('Uninstall complete.');
});

//防止程序运行2次
svc.on('alreadyinstalled', () => {
    console.log('This service is already installed.');
})

//存在就卸载
if(svc.exists) return svc.uninstall();
//不存在就安装
svc.install()