打开html默认生成3个不同的div,输入以下可以继续生成

var div = document.createElement("div");
div.style.width ="300px";
div.style.height = "300px";
div.style.border = "2px solid green";
document.body.appendChild(div);
let extraFunctionObj={selectExcel:true,resizeExcel:true,insertOrDeleteExcel:true};
new Excel(div, 20, 10,extraFunctionObj);

分为一下四个模块 初始化时参数参数如下
 constructor(host, rowNumber, columnNumber,extraFunctionObj) 
host;//生成目标区域 必须 div element 类型
rowNumber; //行数 必须	number类型
columnNumber;//列数 必须 number类型
extraFunctionObj//对象
selectExcel;//是否添加select功能 非必须 boolean
resizeExcel;//是否添加resize功能 非必须 boolean
insertOrDeleteExcel;//是否添加insert or delete 功能 非必须 boolean

1 基本功能 绘制title input target 创建静态页面  创建时一定产生的
2 点击target 选中效果 拖拽效果 双击target input  textchange   可选 
4 resize 点击title	可选
5 add/delete		可选

V9.0：
1.修改了css类选择器命名 统一为以 "_" 连接。
2.将contentList修改为contentArr二维数组并重写相应函数，读取时不用再次遍历。
3.删除menuList的opacity以及disable属性，改用display：none。
4.清除了作用于整个DOM的事件，该为作用于每个host。
5.删除了不必要的注释，添加了函数注释。