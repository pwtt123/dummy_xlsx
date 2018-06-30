 var dummy_xlsx=require("../dummy_xlsx.js");
 var fs=require("fs");

 try {fs.mkdirSync("./export")}catch(err){};

 dummy_xlsx.addConfigsFromFiles("./profile.json");
 console.log("configList",dummy_xlsx.configList);

 dummy_xlsx
  //开始处理模板
     .on("templateStarted",(xPath)=>{
     console.log("开始处理模板：", xPath);
   })
     // 字典加载
     .on("dictLoaded",(xPath)=>{
         console.log("字典已加载：", xPath);

     })
     // dummy 数据 文件 创建成功
     .on("dummyFileCreated",(xPath)=>{
         console.log("生成dummy文件：",xPath)

     })
     //结束处理模板
     .on("templateFinished",(xPath)=>{
         console.log("模板处理完成：", xPath);

     })
     // 出错

     .on("error",(err)=>{
        console.log("err",err)
     });


 // 开始 生成 数据 文件
 dummy_xlsx.doDummy();


