## 概述

此模块用于快速批量生成 随机Dummy数据的 xlsx文件

此模块 需要输入，

* 模板xlsx路径，
* 字典文件路径，
* 输出Dummy xlsx文件的路径，数量，sheet属性



## Getting start

Install with npm:

```
npm install dummy_xlsx
```

Then `require` and use it in your code:

每一个 config 对象参数

* "templateFilePath":"./订单表_模板.xlsx",//源文件路径

* "dictPaths":["./订单表dict.json","./财务报表dict.json"],//字典文件路径集合

* "options":{

  * exportFilesNum":5, //导出文件数量" 

  * "exportFilePaths":["./订单表1.xlsx","./订单表2.xlsx"],//导出文件路径集合（可缺省）

  * "exportFileSheets":{//导出文件sheet属性：

    * "订单表":{

      * "toSheets":["订单1","订单2","订单3"], // 目标sheet
      * "pick":{"min":1,"max":3}，// 随机sheet数

      }

```
var dummy_xlsx=require("dummy._xlsx");

// 添加配置文件
 dummy_xlsx.addConfigsFromFiles("./profile.json");

或者 直接添加 config
addConfigs({...})

// 开始 生成 数据 文件
dummy_xlsx.doDummy();

```



**listener**

```
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
```





## 架构

### 模块 I/O

输入：配置json文件 

路径输出：生成的Dummy数据 xlsx文件

### 配置json文件

#### 定义：

存放 批量生成数据 的输入参数 的json文件

#### 文件内容：

array的json格式，

* array每一项对应一个输入模板文件配置 [ 模板文件配置1，模板文件配置2 , ... ]
  * 模板文件配置对象 ：{...}
    * templateFilePath，输入模板文件路径,
    * dictFilePaths:[]，输入字典文件路径集合,                     
      加载多个字典，                    
      相同属性后加载的字典会覆盖先加载的
    * options{...}，生成文件选项，
      * exportFilesNum，生成Dummy文件个数，
      * exportFilePaths:[]，生成Dummy文件路径集合， 未指定则使用  <源文件路径+复制次数下标> 作为默认路径，若此集合长度小于exportFilesNum，剩余文件路径使用默认路径
      * exportFileSheets:{<模板xlsx的sheet名1>：模板xlsx的sheet对象1,...}，
        * 生成的Dummy数据xlsx 的sheet名以及来源，模板xlsx的sheet对象：{...}
          * toSheets:[]，输出xlsx中 sheet名集合，
          * pick:{min:<最小pick数>,max:<最大pick数>}，数组中随机取值范围，
            无此属性则取全部
            没有此属性，则数据xlsx 和模板xlsx sheet 一一对应    

### 字典文件

#### 定义：

存放 每种字典属性的取值范围 数据的json文件， 

#### 文件内容：

obj的json格式，每一项代表一个字典属性的取值范围，

key 为 字典属性名称，value 如果是数组，则会从数组中pick 一项作为取值，如果是字符串，则会认为是代码，执

行此字符串，结果作为取值{  <字典属性名>:["可能值1"，"可能值2"...],
<字典属性名>:"chance.integer({min:10,max:100})", ...}

### 模板xlsx文件

#### 定义：

指定输出xlsx的数据格式的xlsx文件，在此模板xlsx文件中，可以通过填写 含有【自定义\$\$批注】或【自定义\$\$公式】的单元格，指定输出xlsx文件的数据格式，

#### 模板设置：

* 单元格批注
  * 非自定义\$\$批注，输出文件中不会包含这些批注
  * 自定义\$\$批注
    * \$\$repeatStart：{min:N，max:N}，指定 复制区域的开头，以及需要复制的 次数，同一repeat区域内相同字典值也相同，每个repeatStart,开始，会重新声明一遍所有字典变量，同一repeat区域内同一中字典属性的值不会改变，
    * \$\$repeatEnd，指定复制区域的结尾
    * $$comment，忽略其中一行，输出文件中不会包含此行
    * $$copyUntil:N，从出现此行，一直到指定的行，每行都复制填充此行列的内容，主要用于用于固定表格行的样式
    * $$fixed:line，固定此行位置，模板中在第几行，生成的数据就在第几行，不会随着循环而递增行号


* 单元格内容
  * excel 原有公式，不会保留公式，只会保留计算出的结果
  * 自定义\$\$公式（且无 excel公式），判断条件：单元格内容 含有 \$\$ 字符对于此类单元格，系统会执行eval(<单元格内的内容>) 在输出文件中替换为模板单元格的内容值及格式 （文本，时间，数字），模板单元格内容可能包括：
    * \$\$xx，xx为字典属性名， 输出为最后生效的字典属性定义范围内的随机值。
    * \$\$_chance，这是一个保留的chancejs 对象，可以使用$$_chance.xxx(xxx) 调用chancejs的方法生成随机值，
    * $$_dict，这是一个保留对象，存放 当前sheet 所有字典属性 按声明顺序 依次 声明过的值（每次repeat循环会重新声明一次），
      {<字典属性名>:[<此字典属性第1次随机值>,<此字典属性第2次随机值>,...]}
      可以使用此属性计算某个字典属性的合计等信息 ，
      例：  \_.max($$_dict["\$\$注册资金"])
    * 可以是上述内容的运算组合，算符包括：+,-,. 等，eval支持的 js运算符。
      如：\$\$姓+\$\$名, \$\$family.cnName，$$_chance.guid()
      可以使用 undersoce，增加了 \_.sum(list) 求合计 ，\_.avg(list)求平均数 两个方法

### 生成的Dummy数据 xlsx文件

#### 定义：

按特定的格式生成的Dummy数据的xlsx文件 

## 示例  

**具体参照**

./sample/source.xlsx ，模板xlsx

./sapmple/dicts ，字典文件

./sapmple/profile.json，config属性的配置文件

./sapmple/export，生成的数据文件输出路径

双击 start.cmd 查看测试结果



profile.json:

```
[{
//源文件路径
"templateFilePath":"./订单表_模板.xlsx",
//字典文件路径
"dictPaths":["./订单表dict.json","./财务报表dict.json"],
"options":{  
		//导出文件数量  
		"exportFilesNum":5,
		//导出文件路径集合（可缺省）
        "exportFilePaths":["./订单表1.xlsx","./订单表2.xlsx"],
        //导出文件sheet属性：
        "exportFileSheets":{"订单表":
        					//{<原sheet名>:{toSheets:[<目标sheet名集合>]}}
        					{"toSheets":["订单1","订单2","订单3"],                                    						// 随机sheet数
        					"pick":{"min":1,"max":3}                                
							}
        }
}},
{"templateFilePath":"./财务报表_模板.xlsx",
"dictPaths":["./订单表dict.json","./财务报表dict.json"],
"options":{    
		"exportFilesNum":5,
        "exportFilePaths":["./财务报表1.xlsx","./财务报表2.xlsx"],
        "exportFileSheets":{"财务报表":{"toSheets":["报表1","报表2","报表3"],                                    "pick":{"min":1,"max":3}                                
							}
          }
}}
]
```

订单表dict.json:

```
{"$$编号":"chance.integer({min:1,max:1000})",
"$$交易日期":"chance.date({year:2018})",
"$$交易额":"chance.integer({min:10000,max:100000})",
"$$商品名1":["牙刷","毛巾","肥皂"],
"$$供应商1":["联华超市","苏果超市"]
}
```

财务报表dict.json:

```
{"$$年度销售额":"chance.integer({min:100000,max:1000000})",
"$$本月销售额":"chance.integer({min:10000,max:100000})",
"$$年度活动收入额":"chance.integer({min:10000,max:100000})",
"$$本月活动收入额":"chance.integer({min:10000,max:100000})",
"$$年度税率支出额":"chance.integer({min:10000,max:100000})",
"$$本月税率支出额":"chance.integer({min:10000,max:100000})",
"$$年度其他收入额":"chance.integer({min:10000,max:100000})",
"$$本月其他收入额":"chance.integer({min:10000,max:100000})",
"$$年度薪酬额":"chance.integer({min:10000,max:100000})",
"$$本月薪酬额":"chance.integer({min:10000,max:100000})",
}
```

模板xlsx:

参照 ./sample/source.xlsx