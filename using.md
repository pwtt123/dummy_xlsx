## 概述

此模块用于快速批量生成 Dummy数据的 xlsx文件，

此模块 需要通过 配置文件，指定模板xlsx路径，字典文件路径，以及输出Dummy文件的路径，数量，sheet名等属性，

在模板xlsx中填写特殊批注/内容指定数据来源

在命令行输入配置文件路径，执行即可批量生成Dummy数据xlsx文件

## 流程

* 新建/编辑字典文件，指定各个字典属性的取值范围（详见 设计文档）
* 新建/编辑 模板xlsx 
  * \$\$repeatstart，\$\$repeatend 特殊批注指定需循环的单元格
  * 设定单元格内容及样式，单元格内容可用\$\$<字典属性>、\$\$\_chance、\$\$\_dict 来指定生成数据（详见 设计文档）
* 新建/编辑配置文件，指定模板所使用的字典文件， 生成dummy文件的数量，路径，以及sheet名 和来源（详见 设计文档）
* 新建/编辑 start.cmd 指定配置文件路径，
* 执行start.cmd，生成dummy文件



## 示例

start.cmd：

```
set  PROFILE_PATH = "./profile.json"
node dummy._xlsxjs
```

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

./财务报表_模板.xlsx

./订单表_模板.xlsx