const fs = require("fs");
const _ = require("underscore");
const path = require("path");
const Thenjs = require("thenjs");
const XLSX = require("xlsx");
const XLSX_STYLE = require("xlsx-style");
const Chance = require("chance");
const chance = new Chance();
const $$_chance = chance;

const zRepeatStartMark = "$$repeatstart";
const zRepeatEndMark = "$$repeatend";
const zCommentMark = "$$comment";
const zFixedMark = "$$fixed";
const zCopyUntilMark = "$$copyuntil";

const zProFilePath = process.env.PROFILE_PATH || "";


/*
 each config:
 {
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
 "exportFileSheets":{
 "订单表":{
 // 目标sheet
 "toSheets":["订单1","订单2","订单3"],
 // 随机sheet数
 "pick":{"min":1,"max":3}
 }
 }
 }}
 */


module.exports = {
    configList: [],
    listenerList:{
        //开始处理模板
        templateStarted(xPath){
            // console.log("开始处理模板：", xPath);
        },
        // 字典加载
        dictLoaded(xPath){
            // console.log("字典已加载：", xPath);
        },
        // dummy 数据 文件 创建成功
        dummyFileCreated(xPath){
            // console.log("生成dummy文件：",xPath)
        },
        //结束处理模板
        templateFinished(xPath){
            // console.log("模板处理万郴：", xPath);
        },

        // 出错
        error(err){
            // console.log("err",err);
        }
    },
    // 增加 一个或多个 config 对象
    addConfigs(xConfigs){
        let zConfigs = xConfigs;
        if (!_.isArray(xConfigs)) {
            zConfigs = [xConfigs]
        }
        // 保存到内存
        _.each(zConfigs, (xxConfig) => {
            this.configList.push(xxConfig)
        })
    },
    // 根据 一个或多个 配置文件 ，增加 一个或多个config 对象
    addConfigsFromFiles(xConfigPaths){
        let zConfigPaths = xConfigPaths;
        if (!_.isArray(xConfigPaths)) zConfigPaths = [xConfigPaths];
        try {
            _.each(zConfigPaths, (xxConfigPath) => {
                let zConfigData = fs.readFileSync(xxConfigPath, "utf8");
                let zStartString = "";
                let zIndexA = zConfigData.indexOf("[");
                let zIndexO = zConfigData.indexOf("{");
                zStartString = (zIndexA == -1 ? "{" : zIndexO == -1 ? "[" : zIndexA < zIndexO ? "[" : "{" );

                // parse json 字符的 configs
                let zParseData = JSON.parse(zConfigData.substr(zConfigData.indexOf(zStartString)));
                if (!_.isArray(zParseData)) zParseData = [zParseData];
                // 保存到内存
                _.each(zParseData, (value) => {
                    this.configList.push(value)
                });
            })

        } catch (err) {
            throw Error("configFile parse Error")
        }
    },

    // 删除已存在的 所有config 属性
    clearConfigList(){
        _.each(this.configList,(xValue,xIndex)=>{
            delete this.configList[xIndex]
        })
    },
    /*
     添加监听事件
     //开始处理模板
     templateStarted(xPath)
     // 字典加载
     dictLoaded(xPath)
     // dummy 数据 文件 创建成功
     dummyFileCreated(xPath)
     //结束处理模板
     templateFinished(xPath)
     // 出错
     error(err)
     xCallBack:回调
     */
    on(xEventName,xCallBack){
        if(this.listenerList[xEventName])this.listenerList[xEventName]=xCallBack;
        return this
    },
    off(xEventName){
        if(this.listenerList[xEventName])this.listenerList[xEventName]=()=>{};
        return this;
    },

    // 每一个 模板的 数据生成
    doDummy(){
        // configuration
        let zProOptionList = this.configList;
        // read configuration file
        Thenjs
        // handle each template
            .eachLimit(zProOptionList, (cont, xProOption, xx) => {

                // console.log("zProOptionList", zProOptionList)

                //  dict list，每个字典属性的候选值/代码
                let zDicts = {};
                //读取所有字典
                Thenjs.eachLimit(xProOption.dictFilePaths, (templateCont, xDictPath) => {
                    fs.readFile(xDictPath, "utf8", (err, xDictContent) => {
                        if (err)return cont(err);
                        try {
                            if (xDictContent.indexOf("{") < 0)return templateCont("templateContentIllegal:" + xDictPath);

                            // console.log("aa",xDictContent.indexOf("{"))
                            // console.log("bb",xDictContent.substr(xDictContent.indexOf("{")))
                            // console.log("cc",xDictContent.substr(xDictContent.indexOf("{")).indexOf("{"))

                            let zDictContent = xDictContent.substring(xDictContent.indexOf("{"));

                            _.extend(zDicts, JSON.parse(zDictContent));

                            this.listenerList.dictLoaded(xDictPath);
                            templateCont();

                        } catch (err) {
                            this.listenerList.error(err);
                            // console.log("loadDictErr", xDictPath);
                            templateCont(err);
                        }
                    })
                }, 10)
                //读取模板
                    .then((templateCont) => {
                        this.listenerList.templateStarted(xProOption.templateFilePath);
                        fs.readFile(xProOption.templateFilePath, templateCont);
                    })
                    // 处理模板数据
                    .then((templateCont, xTemplateContent) => {
                        let workbook = XLSX_STYLE.read(xTemplateContent, {
                            type: 'buffer',
                            cellDates: true,
                            cellStyles: true,
                            WTF: false
                        });


                        //单元格 按 先列后行，排序之后的sheet数组集合
                        let zSortedSheetList = {};

                        _.each(workbook["Sheets"], (xSheet, xSheetName) => {
                            if (!zSortedSheetList[xSheetName]) zSortedSheetList[xSheetName] = [];
                            _.each(xSheet, (xCellObj, xCellKey) => {
                                if (xCellKey.indexOf("!") > -1)return;

                                zSortedSheetList[xSheetName].push({
                                    cellKey: xCellKey,
                                    cellObj: xCellObj,
                                    sortIndex: reverseLetterNumber(xCellKey, 6)
                                })
                            });
                            zSortedSheetList[xSheetName] = _.sortBy(zSortedSheetList[xSheetName], "sortIndex")
                        });


                        // console.log("zSortedSheetList",zSortedSheetList);

                        // 生成文件列表
                        let zToFilePaths = [];
                        // 生成文件数
                        let zExportFilesNum = Number(xProOption.options.exportFilesNum) || 0;

                        // console.log("zExportFilesNum",xProOption.options)

                        for (let j = 0; j < zExportFilesNum; j++) {
                            let zParsedPath = path.parse(xProOption.templateFilePath);
                            let zToFilePath = path.join(zParsedPath.dir, zParsedPath.name + "-" + (j + 1) + zParsedPath.ext);
                            if (xProOption.options.exportFilePaths && xProOption.options.exportFilePaths[j]) {
                                zToFilePath = xProOption.options.exportFilePaths[j]
                            }
                            zToFilePaths.push(zToFilePath)
                        }

                        // console.log("zToFilePaths",zToFilePaths)

                        // 循环每个待生成文件
                        _.each(zToFilePaths, (xToFilePath) => {


                            var wb = XLSX.utils.book_new();
                            //处理 每个用到的模板sheet
                            _.each(xProOption.options.exportFileSheets || {}, (xToSheetsObj, xSourceSName) => {

                                // console.log("aa",xToSheetsObj)
                                // toSheets 取值个数
                                var zPickNum = xToSheetsObj.toSheets.length;
                                if (xToSheetsObj.pick) zPickNum = chance.integer(xToSheetsObj.pick);
                                // 生成的sheet名列表
                                var zToSheets = chance.pickset(_.uniq(xToSheetsObj.toSheets), zPickNum);


                                // copy的行集合
                                let zCopyRowList = {};
                                // 最新 copy 的行号
                                let zCopyRow = -1;
                                let zCopyUntilRow = -1;
                                // 是否已拷贝
                                let zIfCopied = true;
                                //最新 fixed 的行号
                                let zFixRow = -1;
                                // 处理每个生成的sheet
                                _.each(zToSheets, (xTSName) => {

                                    //保存每个字典属性使用过的值
                                    let $$_dict = {};

                                    if (!workbook["Sheets"][xSourceSName] || !workbook["Sheets"][xSourceSName]["!ref"] || !zSortedSheetList[xSourceSName]) {
                                        console.log("sheetIsNull", xSourceSName, workbook["Sheets"], zSortedSheetList);
                                        return;
                                    }

                                    let zSourceRef = workbook["Sheets"][xSourceSName]["!ref"];
                                    let zSourceCols = workbook["Sheets"][xSourceSName]["!cols"];
                                    // 待写入的sheet
                                    let zWS2Write = {};
                                    //有效起始单元格
                                    let zStartCell = zSourceRef.split(":")[0];
                                    //有效结束单元格
                                    let zEndCell = zSourceRef.split(":")[1];

                                    // console.log("zSortedSheetList[xSourceSName]",zSortedSheetList[xSourceSName]);

                                    // repeatStart过的key 集合,遍历时只有第一次生效
                                    let zStartedKeys = {};
                                    // 是否重新声明
                                    let zIfReDeclare = true;
                                    //遍历开始的 下标
                                    let zLoopBeginIndex = -1;
                                    // 遍历次数
                                    let zLoopTimes = 0;
                                    //增加的列数
                                    let zRowNum2Add = 0;
                                    let getEvalCell = (v) => {
                                        return v
                                    };

                                    let zCommentRow = 0;

                                    //遍历 原sheet的每个单元格
                                    for (let i = 0; i < zSortedSheetList[xSourceSName].length; i++) {
                                        let zThisCell = zSortedSheetList[xSourceSName][i];

                                        // console.log("i",i)
                                        // 单元格地址
                                        let zCellLocation = cutLetterNumber(zThisCell.cellKey);

                                        let zIfNeedAddRow = false;

                                        // 若有批注，保存特殊批注
                                        if (zThisCell["cellObj"]["c"]) {

                                            // 处理每个用户的备注
                                            _.each(zThisCell["cellObj"]["c"], function (xComment, xCommentIndex) {
                                                if (!xComment || !xComment["t"] || typeof (xComment["t"]) != "string")return;
                                                // 按照换行符分割
                                                const zSplitedComment = xComment["t"].split("\n");
                                                // 处理每个行 备注
                                                _.each(zSplitedComment, function (xxComment) {
                                                    xxComment = xxComment.toLowerCase();
                                                    // console.log("xxComment",xxComment)
                                                    const zIfRepeatStart = (xxComment.indexOf(zRepeatStartMark) > -1);
                                                    const zIfRepeatEnd = (xxComment.indexOf(zRepeatEndMark) > -1);
                                                    const zIfComment = (xxComment.indexOf(zCommentMark) > -1);
                                                    const zIfFixed = (xxComment.indexOf(zFixedMark) > -1);
                                                    const zIfCopy = (xxComment.indexOf(zCopyUntilMark) > -1);

                                                    // 未匹配到 特殊批注,返回
                                                    if (!zIfRepeatStart && !zIfRepeatEnd && !zIfComment && !zIfFixed && !zIfCopy) {
                                                        return;
                                                    }
                                                    // 匹配到特殊批注，保存此单元格批注
                                                    else {
                                                        //保存 特殊标记信息 信息
                                                        if (zIfRepeatStart) {
                                                            // 若已经读取过此批注，返回
                                                            if (zStartedKeys[zThisCell.cellKey])return;
                                                            zStartedKeys[zThisCell.cellKey] = 1;
                                                            // 去掉 $$key,:： 和 \n,剩余字符去除前后空格作为 自定义的列头
                                                            let zIntegerObj = {min: 1, max: 1};
                                                            let zRepeatStartValue = xxComment.replace(zRepeatStartMark, "").replace(/[:：]/, "").replace(/\n/, "").trim();
                                                            try {
                                                                eval("zIntegerObj = " + zRepeatStartValue)
                                                            } catch (err) {
                                                            }
                                                            ;

                                                            zLoopBeginIndex = i;
                                                            zLoopTimes = chance.integer(zIntegerObj);
                                                            // console.log("start",i,"zLoopTimes",zLoopTimes);
                                                        }

                                                        if (zIfRepeatEnd) {
                                                            // 未指定开始 的下标，返回
                                                            if (zLoopBeginIndex < 0)return;
                                                            // console.log("end",i);
                                                            zLoopTimes--;
                                                            // 剩余循环次数大于0，指针跳转到 开始循环位置
                                                            if (zLoopTimes > 0) {
                                                                i = zLoopBeginIndex - 1;
                                                                zIfNeedAddRow = true;
                                                            }
                                                            // delete zThisCell["cellObj"]["c"][xCommentIndex]
                                                        }

                                                        // comment 标记
                                                        if (zIfComment) {

                                                            zCommentRow = Number(zCellLocation.num);
                                                            // 增加行数-1
                                                            zRowNum2Add = zRowNum2Add - 1;

                                                        }

                                                        // 若为 copyUntil 标记
                                                        if (zIfCopy) {
                                                            zCopyRowList = {};
                                                            let zCopyValue = xxComment.replace(zCopyUntilMark, "").replace(/[:：]/, "").replace(/\n/, "").trim();
                                                            zCopyRow = Number(zCellLocation.num) + zRowNum2Add;
                                                            zCopyUntilRow = Number(zCopyValue);
                                                            zIfCopied = false;
                                                        }

                                                        // fixed 标记
                                                        if (zIfFixed) {
                                                            let zFixedValue = xxComment.replace(zFixedMark, "").replace(/[:：]/, "").replace(/\n/, "").trim();
                                                            // console.log("aa",zFixedValue)

                                                            if (zFixedValue == "line") {
                                                                // console.log("bb")
                                                                zFixRow = Number(zCellLocation.num);
                                                            }
                                                        }

                                                    }
                                                });

                                            });


                                        }
                                        // 若需要重新声明变量
                                        if (zIfReDeclare) {

                                            // 初始化 字典变量
                                            for (let xDictKey in zDicts) {
                                                //$$_开头特殊变量不初始化
                                                if (xDictKey.indexOf("$$_") === 0)return;

                                                try {
                                                    //从字典中取的随机值
                                                    let zCurrentDictValue = "";
                                                    if (typeof zDicts[xDictKey] == "object") zCurrentDictValue = chance.pickone(zDicts[xDictKey]);
                                                    if (typeof zDicts[xDictKey] == "string") zCurrentDictValue = eval(zDicts[xDictKey]);
                                                    // 初始化
                                                    let zEvalString = "var " + xDictKey + " = '';";
                                                    if (typeof zCurrentDictValue == "number") zEvalString = "var " + xDictKey + " = " + zCurrentDictValue + ";";
                                                    else if (typeof zCurrentDictValue == "string") zEvalString = "var " + xDictKey + " = '" + zCurrentDictValue + "';";
                                                    else if (zCurrentDictValue instanceof Date) zEvalString = "var " + xDictKey + " = new Date(" + JSON.stringify(zCurrentDictValue) + ");";
                                                    else if (typeof zCurrentDictValue == "object") zEvalString = "var " + xDictKey + " = '" + JSON.stringify(zCurrentDictValue) + "';";
                                                    eval(zEvalString);

                                                    // let zTest=eval(xDictKey);
                                                    // console.log("初始化",xDictKey,zTest,typeof zTest)

                                                    if (!$$_dict[xDictKey]) $$_dict[xDictKey] = [];
                                                    $$_dict[xDictKey].push(zCurrentDictValue);

                                                    // console.log("dict push",xDictKey,zCurrentDictValue)

                                                } catch (err) {
                                                    this.listenerList.error(err);
                                                    // console.log("字典属性：", xDictKey, "解析错误", err);
                                                }


                                            }

                                            // 获取 eval 后的 cell 对象
                                            getEvalCell = function (xCellObj) {


                                                // console.log("xCellObj",xCellObj);
                                                var zReturnCellObj = {};

                                                _.extend(zReturnCellObj, _.pick(xCellObj, "v", "f", "w", "t", "s"));
                                                if (zReturnCellObj.c) zReturnCellObj.c.hidden = true;
                                                // 获取文本
                                                let zCell = xCellObj.w || "";
                                                //若有 excel公式,不做处理
                                                if (xCellObj["f"]) {
                                                    // zReturnCellObj.w="="+xCellObj["f"];
                                                    // zReturnCellObj.v="="+xCellObj["f"];
                                                    // zReturnCellObj.t="e";
                                                }
                                                //若没有 excel公式
                                                else {

                                                    // 存在 $$ 字符
                                                    if (zCell && zCell.indexOf("$$") > -1) {

                                                        // 是否可 eval
                                                        let zIfEvalable = true;
                                                        try {

                                                            // console.log("eval",xCellObj.w, eval(zCell))

                                                            zCell = eval(zCell)
                                                        } catch (err) {
                                                            zIfEvalable = false
                                                        }


                                                        // 默认为文本型
                                                        zReturnCellObj.t = "s";
                                                        zReturnCellObj.w = zCell;
                                                        // 若执行了 eval
                                                        if (zIfEvalable) {
                                                            // 若为时间类型
                                                            if (zCell instanceof Date) {
                                                                zReturnCellObj.t = "d";
                                                                zReturnCellObj.w = zCell.format("yyyy-MM-dd hh:mm:ss")
                                                            }

                                                            if (typeof zCell == "number") {
                                                                // console.log("number",xCellObj.w,zCell)
                                                                zReturnCellObj.t = "n";
                                                                zReturnCellObj.w = JSON.stringify(zCell)
                                                            }

                                                            if (typeof zCell == "boolean") {
                                                                zReturnCellObj.t = "b";
                                                                zReturnCellObj.w = JSON.stringify(zCell)
                                                            }
                                                        }

                                                        zReturnCellObj.v = zCell;
                                                    }

                                                }


                                                // console.log("w",zReturnCellObj.w,"v",zReturnCellObj.v)
                                                //
                                                // console.log("zReturnCellObj",zReturnCellObj)

                                                return zReturnCellObj;

                                            };
                                            zIfReDeclare = false;
                                        }
                                        // console.log("oldCell",zThisCell["cellObj"])

                                        let zNewKey = zCellLocation.letter + (Number(zCellLocation.num) + zRowNum2Add);
                                        let zNewCell = getEvalCell(zThisCell["cellObj"]);


                                        // console.log("zNewCell",zNewCell)
                                        if (zCommentRow != Number(zCellLocation.num)) {

                                            // 若此行为待拷贝数据，列放到 zCopyRowList 对象中
                                            if (zCopyRow == Number(zCellLocation.num) + zRowNum2Add) {
                                                // console.log("copyrow",zCellLocation)
                                                zCopyRowList[zCellLocation.letter] = zNewCell;
                                            }


                                            // console.log("row index",zCopyUntilRow,Number(zCellLocation.num) + zRowNum2Add);

                                            // 若未拷贝，但超过需 copy行,则先执行拷贝
                                            if (!zIfCopied && (Number(zCellLocation.num) + zRowNum2Add) > zCopyUntilRow) {
                                                // console.log("zCopyRow",zCopyRow,"",zCopyUntilRow)
                                                // 循环所有行需要拷贝的行
                                                for (let k = zCopyRow; k <= zCopyUntilRow; k++) {
                                                    _.each(zCopyRowList, (xCell, xColIndex) => {
                                                        if (xCell) zWS2Write[xColIndex + k] = xCell;
                                                        // console.log(xColIndex+k,zWS2Write[xColIndex+k])
                                                    })
                                                }

                                                // console.log("zWS2Write",zWS2Write)
                                                zIfCopied = true;
                                            }


                                            //若此行需固定，则覆盖原有地址的单元格
                                            if (zFixRow == Number(zCellLocation.num)) {
                                                // if(zWS2Write[zThisCell.cellKey])console.log(zThisCell.cellKey,"单元格被覆盖");
                                                zWS2Write[zThisCell.cellKey] = zNewCell
                                            }
                                            // 此行为 copy 行，不保存
                                            else if (zCopyRow == Number(zCellLocation.num) + zRowNum2Add) {

                                            }
                                            else {
                                                // if(zWS2Write[zNewKey])console.log(zNewKey,"单元格被覆盖");
                                                zWS2Write[zNewKey] = zNewCell;
                                            }
                                        }


                                        // 此单元格为 repeatEnd 时，生成的行数累加
                                        if (zIfNeedAddRow) {
                                            zIfReDeclare = true;
                                            // 增加的行数，为当前结束行号 - 开始行号+1
                                            let zAddRow = 1 + Number(zCellLocation.num) - Number(cutLetterNumber(zSortedSheetList[xSourceSName][zLoopBeginIndex].cellKey).num);

                                            //每次循环结束，需增加行数累加
                                            zRowNum2Add += zAddRow;
                                        }

                                    }
                                    ;

                                    // 若未拷贝，但超过需 copy行,则先执行拷贝
                                    if (!zIfCopied) {
                                        // 循环所有行需要拷贝的行
                                        for (let k = zCopyRow; k <= zCopyUntilRow; k++) {
                                            _.each(zCopyRowList, (xCell, xColIndex) => {
                                                if (xCell) zWS2Write[xColIndex + k] = xCell;
                                            })
                                        }

                                        zIfCopied = true;
                                    }


                                    let zNewEndRow = Number(cutLetterNumber(zEndCell).num) + zRowNum2Add;
                                    // console.log("zNewEndRow",zNewEndRow)

                                    // 设置特殊属性
                                    zWS2Write["!ref"] = zStartCell + ":" + cutLetterNumber(zEndCell).letter + zNewEndRow;
                                    zWS2Write["!cols"] = zSourceCols;
                                    // console.log("!ref",zWS2Write["!ref"])
                                    // console.log("!cols",zWS2Write["!cols"])

                                    // console.log("zWS2Write",zWS2Write);
                                    XLSX.utils.book_append_sheet(wb, zWS2Write, xTSName)
                                });


                            });

                            XLSX_STYLE.writeFile(wb, xToFilePath);
                            this.listenerList.dummyFileCreated(xToFilePath);

                            // console.log("生成文件：", xToFilePath);
                        });


                        templateCont();
                    })
                    .then((templateCont) => {
                        this.listenerList.templateFinished(xProOption.templateFilePath);
                        // console.log("此模板数据全部生成完毕：", xProOption.templateFilePath);

                        cont();
                        templateCont();
                    })
                    .fail((templateCont, err) => {
                        cont(err);
                    })

            }, 1)
            .then((cont) => {
                // 清空 配置信息
                this.clearConfigList();
                cont()
            })
            .fail((cont, err) => {
                this.listenerList.error(err);
                // console.log("err", err)
            });
    }


};


//传入字母与数字组合字符串：“A1”,分割成字母和数字返回：｛num:1，letter:"A"｝
function cutLetterNumber(xKey) {
    let zTestLetter = /[A-Z]+/;
    let zTestNumber = /[0-9]+/;
    //console.log("zTestLetter",xString.match(zTestLetter));
    //console.log("zTestNumber",zTestNumber.exec(xString));
    let zNum = xKey.match(zTestNumber) ? xKey.match(zTestNumber)[0] : "";
    let zLetter = xKey.match(zTestLetter) ? xKey.match(zTestLetter)[0] : "";
    let zReturnObj = {"num": Number(zNum), "letter": zLetter};
    //this.testLetterNumber['out']=zReturnObj;
    return zReturnObj
}

// 反转 字母和数字 连接的字符串，数字补全位数
function reverseLetterNumber(xKey, xPlaces) {
    var zLocation = cutLetterNumber(xKey);
    return lz(zLocation.num, xPlaces || 0) + lz(zLocation.letter, xPlaces || 0)
}

// (123,4) => "0123"
function lz(num, places) {
    var zero = places - num.toString().length + 1;
    return Array(+(zero > 0 && zero)).join("0") + num;
}

//求合计值
_.sum = function (xList) {
    return _.reduce(xList, (m, v) => {
        return m + Number(v)
    }, 0);
};

//求平均值
_.avg = function (xList) {
    return _.sum(xList) / _.keys(xList).length;
};


Date.prototype.format = function (format) {
    var o = {
        "M+": this.getMonth() + 1, //month
        "d+": this.getDate(),    //day
        "h+": this.getHours(),   //hour
        "m+": this.getMinutes(), //minute
        "s+": this.getSeconds(), //second
        "q+": Math.floor((this.getMonth() + 3) / 3),  //quarter
        "S": this.getMilliseconds() //millisecond
    };
    if (/(y+)/.test(format)) format = format.replace(RegExp.$1,
        (this.getFullYear() + "").substr(4 - RegExp.$1.length));
    for (var k in o)if (new RegExp("(" + k + ")").test(format))
        format = format.replace(RegExp.$1,
            RegExp.$1.length == 1 ? o[k] :
                ("00" + o[k]).substr(("" + o[k]).length));
    return format;
};