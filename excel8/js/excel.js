class Excel {
    constructor(host, rowNumber, columnNumber,extraFunctionObj) {
        this.host = host;//生成目标区域
        this.rowNumber = rowNumber; //行数
        this.columnNumber = columnNumber;//列数
        
        this.selectExcel=extraFunctionObj.selectExcel;//是否添加select功能
        this.resizeExcel=extraFunctionObj.resizeExcel;//是否添加resize功能
        this.insertOrDeleteExcel=extraFunctionObj.insertOrDeleteExcel;//是否添加insert or delete 功能
        this.heightArr = this.createHeightArr(rowNumber)//长度等于行数+1(包含了表头 title)每一行的高度
        this.widthArr = this.createWidthArr(columnNumber);//长度等于列数+1 每一列的宽度
        this.initialize();
        this.createHeader();
        this.createBody();
        this.createLittleDivTitle();
        this.createLittleDivInput();
        this.createLittleDivTarget();
        this.extraFunction();
    }
    initialize(){
        this.input_display={};
        this.menuListRoot={};
        this.excelTitleRoot={};
        this.excelInputRoot={};
        this.excelTargetRoot={};
        this.titleDivOfX=[{}];
        this.titleDivOfY=[{}];
        this.inputDiv={};
        this.targetDiv=[{}];
        this.insertButton={};
        this.deleteButton={};
        this.overflowBody={};
        this.selectTitle = {};
        this.selectTitleID = "";
        this.selected = {};
        this.isSelected = false;
        this.titleDownID="";
        this.titleDowned={};
        this.isDown = false;
        this.downID = "";  
        this.leftID = 0;
        this.topID = 0;
        this.rightID = 0;
        this.bottomID = 0;
        this.marginTopPix=0;
        this.isDblClicked = false;
        this.contentList = [];// 稀疏矩阵 行_列_内容
        
    }
    createHeightArr(rowNumber) {
        let heightArr = [20];
        for (let i = 0; i < rowNumber; i++) {
            heightArr.push(20);
        }
        return heightArr;
    }
    createWidthArr(columnNumber) {
        let widthArr = [20];
        for (let i = 0; i < columnNumber; i++) {
            widthArr.push(60);
        }
        return widthArr;
    }
    createHeader() {
        let node = document.createDocumentFragment();

        let header = document.createElement("div");
        header.className = "div-header";
        let headerSpan = document.createElement("span");
        headerSpan.className = "div-header-span";
        let displayInput = document.createElement("input");
        displayInput.className = "input-display";
        displayInput.id = "input-display";
        displayInput.setAttribute("disabled", "disabled");
        displayInput.setAttribute("type", "text");
        this.input_display=displayInput;
        node.appendChild(displayInput);
        headerSpan.appendChild(node);
        header.appendChild(headerSpan);
        this.host.appendChild(header);
        this.host.style.userSelect = "none";
        //选择功能 添加事件
        document.removeEventListener("mouseup",this.mouseUp);
        document.addEventListener("mouseup", this.mouseUp.bind(this));
        document.addEventListener("contextmenu",function(event){
            event.preventDefault();
        });
        this.marginTopPix=Number.parseInt(this.host.style.height)/10+this.host.offsetTop;
    }
    createBody() {
        let node = document.createDocumentFragment();
        let body = document.createElement("div");
        body.className = "div-body";

        let menuList = document.createElement("div");
        menuList.id = "menuList";
        menuList.className = "menuList";

        let insertMenu = document.createElement("button");
        
        let insertIcon=document.createElement("img");
        insertIcon.src="img/insertIcon.png";
        insertIcon.className="insertIcon";

        insertMenu.className = "menuButton";
        insertMenu.id = "insertMenu";
        insertMenu.innerHTML = "Insert";

        insertMenu.appendChild(insertIcon);
        this.insertButton=insertMenu;
        menuList.appendChild(insertMenu);

        let deleteMenu = document.createElement("button");
        
        let deleteIcon=document.createElement("img");
        deleteIcon.src="img/deleteIcon.png";
        deleteIcon.className="deleteIcon";
        

        deleteMenu.className = "menuButton";
        deleteMenu.id = "deleteMenu";
        deleteMenu.innerHTML = "Delete";

        deleteMenu.appendChild(deleteIcon);
        this.deleteButton=deleteMenu;
        menuList.appendChild(deleteMenu);

        this.menuListRoot=menuList;

        node.appendChild(menuList);
        

        let excelTitle = document.createElement("div");
        excelTitle.className = "div-body-title";
        excelTitle.id = "excelTitle";
        this.excelTitleRoot=excelTitle;
        node.appendChild(excelTitle);

        let excelInput = document.createElement("div");
        excelInput.className = "div-body-input";
        excelInput.id = "excelInput";
        this.excelInputRoot=excelInput;
        node.appendChild(excelInput);

        let excelTarget = document.createElement("div");
        excelTarget.className = "div-body-target";
        excelTarget.id = "excelTarget";
        this.excelTargetRoot=excelTarget;
        node.appendChild(excelTarget);

        this.overflowBody=body;
        body.appendChild(node);
        this.host.appendChild(body);
    }

    createLittleDivTitle() {
        let node = document.createDocumentFragment();
        let leftPix = -this.widthArr[0];
        let count = 1;
        let xCode1 = 65;
        let xCode2 = 65;
        let yCode = 1;
        this.titleDivOfY=[{}];
        this.titleDivOfX=[{}];
        for (let j = 0; j <= this.columnNumber; j++) {
            let littleDivObj = document.createElement("div");
            littleDivObj.style.width = this.widthArr[j] + "px";
            littleDivObj.style.height = this.heightArr[0] + "px";
            littleDivObj.style.marginLeft = ((leftPix += this.widthArr[(j - 1) < 0 ? 0 : (j - 1)]) + j) + "px";
            littleDivObj.className = "littleDivTitleX";
            littleDivObj.id = "div0_" + j;
            if (count == 1) {
                count++;
                this.titleDivOfY.push(littleDivObj);/////////////////////////////////////////////////////////////////////////////////
                node.appendChild(littleDivObj);
                continue;
            }
            count++;
            if (count > 28) {
                xCode1 = (xCode1 - 65) % 26 + 65;
                littleDivObj.innerHTML = String.fromCharCode((xCode2 + Number(count - 29) / 26), xCode1++);
            } else {
                littleDivObj.innerHTML = String.fromCharCode(xCode1++);
            }
            this.titleDivOfY.push(littleDivObj);/////////////////////////////////////////////////////////////////////////////////
            node.appendChild(littleDivObj);
        }
        let topPix = -this.heightArr[0];
        count = 1;
        for (let i = 0; i <= this.rowNumber; i++) {
            let littleDivObj = document.createElement("div");
            littleDivObj.style.width = this.widthArr[0] + "px";
            littleDivObj.style.height = this.heightArr[i] + "px";
            littleDivObj.style.marginTop = ((topPix += this.heightArr[(i - 1) < 0 ? 0 : (i - 1)]) + i) + "px";
            littleDivObj.className = "littleDivTitleY";
            littleDivObj.id = "div" + i + "_0";
            if (count == 1) {
                count++;
                this.titleDivOfX.push(littleDivObj);/////////////////////////////////////////////////////////////////////////////////
                node.appendChild(littleDivObj);
                continue;
            }
            littleDivObj.innerHTML = yCode++;
            this.titleDivOfX.push(littleDivObj);/////////////////////////////////////////////////////////////////////////////////
            node.appendChild(littleDivObj);
        }
        this.excelTitleRoot.appendChild(node);
    }

    createLittleDivInput() {
        this.inputDiv={};
        let littleDivObj = document.createElement("input");
        littleDivObj.style.width = this.widthArr[1] + "px";
        littleDivObj.style.height = this.heightArr[1] + "px";
        littleDivObj.style.marginLeft ="21px";
        littleDivObj.style.marginTop ="21px";
        littleDivObj.className = "littleDivInput";
        littleDivObj.id = "inputDiv";
        littleDivObj.setAttribute("disabled", "disabled");
        this.inputDiv=littleDivObj;/////////////////////////////////////////////////////////////////////////////////
        this.excelInputRoot.appendChild(littleDivObj);
    }

    createLittleDivTarget() {
        let node = document.createDocumentFragment();
        let leftPix = 1;
        let topPix = this.heightArr[0]+1;
        this.targetDiv=[{}];
        for (let i = 1; i <= this.rowNumber; i++) {
            for (let j = 1; j <= this.columnNumber; j++) {
                let littleDivObj = document.createElement("div");
                littleDivObj.style.width = this.widthArr[j] + "px";
                littleDivObj.style.height = this.heightArr[i] + "px";
                littleDivObj.style.marginLeft = ((leftPix += this.widthArr[(j - 1)]) + j) + "px";
                littleDivObj.style.marginTop = (topPix + i) + "px";
                littleDivObj.className = "littleDivTarget";
                littleDivObj.id = "div" + i + "_" + j;
                for (let k = 0; k < this.contentList.length; k++) {
                    let contentStr = this.contentList[k].split("_");
                    let rowNum = contentStr[0];
                    let colNum = contentStr[1];
                    if ((i == rowNum) && (j == colNum)) {
                        littleDivObj.innerHTML = contentStr[2];
                        break;
                    }
                }
                this.targetDiv.push(littleDivObj);/////////////////////////////////////////////////////////////////////////////////
                node.appendChild(littleDivObj);


            }
            leftPix = 1;
            topPix += this.heightArr[i];
        }
        this.excelTargetRoot.appendChild(node);
    }

    extraFunction(){
        let that=this;
        if(this.selectExcel){
            let subSelectExcel = new SelectExcel(that);
        }
        if(this.resizeExcel){
            let subResizeExcel = new ResizeExcel(that);
        }
        if(this.insertOrDeleteExcel){
            let subInsertOrDeleteExcel=new InsertOrDeleteExcel(that);
        }

    }
    mouseUp(event) {
        if((this.titleDownID!="")){
            let numY = Number(this.titleDownID.substring(3).split("_")[1]);
            let changedWidth = parseInt(this.titleDowned.style.width);
            changedWidth = (changedWidth < 30) ? 30 : changedWidth;
            this.widthArr[numY] = changedWidth;
            this.widthArr[0] = 20;
            let numX = Number(this.titleDownID.substring(3).split("_")[0]);
            let changedHeight = parseInt(this.titleDowned.style.height);
            changedHeight = (changedHeight < 15) ? 15 : changedHeight;
            this.heightArr[numX] = changedHeight;
            this.heightArr[0] = 20;
            this.clean();
            this.createLittleDivTitle();
            this.titleDownID="";
            this.createLittleDivTarget();
        }
        this.isDown = false;
    }

    clean() {
        //console.log(this);
        let titleNodeXs = this.host.querySelectorAll(".littleDivTitleX");
        for (let titleX of titleNodeXs) {
            titleX.parentNode.removeChild(titleX);
        }
        let titleNodeYs = this.host.querySelectorAll(".littleDivTitleY");
        for (let titleY of titleNodeYs) {
            titleY.parentNode.removeChild(titleY);
        }
        let targetNodes = this.host.querySelectorAll(".littleDivTarget");
        for (let targetNode of targetNodes) {
            targetNode.parentNode.removeChild(targetNode);
        }
    }

    cleanSelect() {
        if (this.isSelected) {
            let clearX = this.titleDivOfX[Number(this.downID.split("_")[0])+1];//document.getElementById("div" + that.downID.split("_")[0] + "_0");
            let clearY = this.titleDivOfY[Number(this.downID.split("_")[1])+1];//document.getElementById("div0_" + that.downID.split("_")[1]);
            clearX.style.borderLeft = "solid 1px #A9A9A9";
            clearY.style.borderTop = "solid 1px #A9A9A9";
            clearX.style.zIndex = 0;
            clearY.style.zIndex = 0;
            clearX.style.color = "black";
            clearY.style.color = "black";
            clearX.style.backgroundColor = "rgba(220,220,220,0.8)";
            clearY.style.backgroundColor = "rgba(220,220,220,0.8)";
            this.selected.style.border = "solid 1px #A9A9A9";
            this.selected.style.zIndex = 0;
        }
        for (let x = this.leftID; x <= this.rightID; x++) {
            for (let y = this.topID; y <= this.bottomID; y++) {
                let pos=(x-1)*this.columnNumber+y;
                if(pos<0){
                    break;
                }
                this.targetDiv[pos].style.backgroundColor = "white";//document.getElementById("div" + x + "_" + y).style.backgroundColor = "white";
            }
        }
        if (this.isDblClicked) {
            let inputObj = this.inputDiv//document.getElementById("inputDiv");
            inputObj.style.zIndex = 0;
            inputObj.style.opacity = 0;
            inputObj.style.margin="21px 0 0 21px";

            inputObj.setAttribute("disabled", "disabled");
        }
        if (this.selectTitleID != "") {
            if (this.selectTitleID.split("_")[0] == 0) {
                this.selectTitle.style.color = "black";
                this.selectTitle.style.backgroundColor = "rgba(220,220,220,0.8)";
                this.selectTitle.style.borderTop = "solid 1px #A9A9A9";
                for (let i = 1; i <= this.rowNumber; i++) {
                    let pos=(i-1)*this.columnNumber+Number(this.selectTitleID.split("_")[1]);
                    this.targetDiv[pos].style.backgroundColor = "white";//document.getElementById("div" + i + "_" + that.selectTitleID.split("_")[1]).style.backgroundColor = "white";
                }
            } else if (this.selectTitleID.split("_")[1] == 0) {
                this.selectTitle.style.color = "black";
                this.selectTitle.style.backgroundColor = "rgba(220,220,220,0.8)";
                this.selectTitle.style.borderLeft = "solid 1px #A9A9A9";
                for (let i = 1; i <= this.columnNumber; i++) {
                    let pos=(Number(this.selectTitleID.split("_")[0])-1)*this.columnNumber+i;
                    this.targetDiv[pos].style.backgroundColor = "white";//document.getElementById("div" + that.selectTitleID.split("_")[0] + "_" + i).style.backgroundColor = "white";
                }
            }
        }
    }
}