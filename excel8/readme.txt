打开html默认生成3个不同的div,输入以下可以继续生成

var div = document.createElement("div");
div.style.width ="300px";
div.style.height = "300px";
div.style.border = "2px solid green";
document.body.appendChild(div);
new Excel(div, 20, 10);

分为一下四个模块 初始化时参数参数如下
 constructor(host, rowNumber, columnNumber,selectExcel,resizeExcel,insertOrDeleteExcel) 
host;//生成目标区域 必须 div element 类型
rowNumber; //行数 必须	number类型
columnNumber;//列数 必须 number类型
selectExcel;//是否添加select功能 非必须 boolean或者数字类型（0为假 其他为真） 默认为否 
resizeExcel;//是否添加resize功能 非必须 boolean或者数字类型（0为假 其他为真） 默认为否 需要添加resize功能时 必须输入select 参数
insertOrDeleteExcel;//是否添加insert or delete 功能 非必须 boolean或者数字类型（0为假 其他为真） 默认为否 需要添加insert or delete功能时 必须输入select 参数以及resize参数。

1 基本功能 绘制title input target  创建时一定产生的
2 点击target 选中效果 拖拽效果 双击target input  textchange   可选
4 resize 点击title	可选
5 add/delete		可选
将方法设置为类 不用静态方法 将传入的参数设置为结构体 方便用户输入调用的方法