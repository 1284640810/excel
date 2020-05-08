class Excel {
    constructor(host, rowNumber, columnNumber, extraFunctionObj) {
        this.host = host;//the host div
        this.rowNumber = rowNumber; 
        this.columnNumber = columnNumber;

        this.selectExcel = extraFunctionObj.selectExcel;//whether add select function 
        this.resizeExcel = extraFunctionObj.resizeExcel;//whether add resize function 
        this.insertOrDeleteExcel = extraFunctionObj.insertOrDeleteExcel;//weather add insortOrDelete function 

        this.heightArr = this.createHeightArr(rowNumber)//height of each row
        this.widthArr = this.createWidthArr(columnNumber);//width of each column
        this.contentArr = this.createContentArr(rowNumber,columnNumber);

        this.initialize();
        this.createHeader();
        this.createBody();
        this.createLittleDivTitle();
        this.createLittleDivInput();
        this.createLittleDivTarget();
        this.extraFunction();
    }
    /**
     * other parameter
     */
    initialize() {
        this.input_display = {};
        this.menuListRoot = {};
        this.excelTitleRoot = {};
        this.excelInputRoot = {};
        this.excelTargetRoot = {};
        this.titleDivOfX = [{}];
        this.titleDivOfY = [{}];
        this.inputDiv = {};
        this.targetDiv = [{}];
        this.insertButton = {};
        this.deleteButton = {};
        this.overflowBody = {};
        this.selectTitle = {};
        this.selectTitleID = "";
        this.selected = {};
        this.isSelected = false;
        this.titleDownID = "";
        this.titleDowned = {};
        this.isDown = false;
        this.downID = "";
        this.leftID = 0;
        this.topID = 0;
        this.rightID = 0;
        this.bottomID = 0;
        this.marginTopPix = 0;
        this.isDblClicked = false;
        
    }
    /**
     * @method createHeightArr
     * @returns {Array} heightArr
     * @param {number} rowNumber 
     * create array of each cell height
     */
    createHeightArr(rowNumber) {
        let heightArr = [20];
        for (let i = 0; i < rowNumber; i++) {
            heightArr.push(20);
        }
        return heightArr;
    }
    /**
     * @method createWidthArr
     * @returns {Array} widthArr
     * @param {number} columnNumber 
     * create array of each cell width
     */
    createWidthArr(columnNumber) {
        let widthArr = [20];
        for (let i = 0; i < columnNumber; i++) {
            widthArr.push(60);
        }
        return widthArr;
    }
    /**
     * @method createContentArr
     * @returns {Array} arr
     * @param {number,number} rowNumber,columnNumber
     * create array to store cell content
     */
    createContentArr(rowNumber,columnNumber){
        let arr=Array(rowNumber+1);
        for(let i=0;i<rowNumber+1;i++){
            arr[i]=Array(columnNumber+1).fill("");
            
        }
        return arr;
    }
    /**
     * @method createHeader 
     * create header DOM
     */
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
        this.input_display = displayInput;
        node.appendChild(displayInput);
        headerSpan.appendChild(node);
        header.appendChild(headerSpan);
        this.host.appendChild(header);
        this.host.style.userSelect = "none";
        //add event 
        document.removeEventListener("mouseup", this.mouseUp);
        document.addEventListener("mouseup", this.mouseUp.bind(this));
        this.host.addEventListener("contextmenu", function (event) {
            event.preventDefault();
        });
        this.marginTopPix = Number.parseInt(this.host.style.height) / 10 + this.host.offsetTop;
    }
    /**
     * @method createHeader 
     * create body DOM
     */
    createBody() {
        let node = document.createDocumentFragment();
        let body = document.createElement("div");
        body.className = "div-body";

        let menuList = document.createElement("div");
        menuList.id = "menuList";
        menuList.className = "div-menu";

        let insertMenu = document.createElement("button");

        let insertIcon = document.createElement("img");
        insertIcon.src = "img/insertIcon.png";
        insertIcon.className = "icon-insert";

        insertMenu.className = "button-menu";
        insertMenu.id = "insertMenu";
        insertMenu.innerHTML = "Insert";

        insertMenu.appendChild(insertIcon);
        this.insertButton = insertMenu;
        menuList.appendChild(insertMenu);

        let deleteMenu = document.createElement("button");

        let deleteIcon = document.createElement("img");
        deleteIcon.src = "img/deleteIcon.png";
        deleteIcon.className = "icon-delete";


        deleteMenu.className = "button-menu";
        deleteMenu.id = "deleteMenu";
        deleteMenu.innerHTML = "Delete";

        deleteMenu.appendChild(deleteIcon);
        this.deleteButton = deleteMenu;
        menuList.appendChild(deleteMenu);

        this.menuListRoot = menuList;

        node.appendChild(menuList);


        let excelTitle = document.createElement("div");
        excelTitle.className = "div-body-title";
        excelTitle.id = "excelTitle";
        this.excelTitleRoot = excelTitle;
        node.appendChild(excelTitle);

        let excelInput = document.createElement("div");
        excelInput.className = "div-body-input";
        excelInput.id = "excelInput";
        this.excelInputRoot = excelInput;
        node.appendChild(excelInput);

        let excelTarget = document.createElement("div");
        excelTarget.className = "div-body-target";
        excelTarget.id = "excelTarget";
        this.excelTargetRoot = excelTarget;
        node.appendChild(excelTarget);

        this.overflowBody = body;
        body.appendChild(node);
        this.host.appendChild(body);
    }

    /**
     * @method createLittleDivTitle 
     * create horizontal and vertical excel title DOM
     */
    createLittleDivTitle() {
        let node = document.createDocumentFragment();
        let leftPix = -this.widthArr[0];
        let count = 1;
        let xCode1 = 65;
        let xCode2 = 65;
        let yCode = 1;
        this.titleDivOfY = [{}];
        this.titleDivOfX = [{}];
        for (let j = 0; j <= this.columnNumber; j++) {
            let littleDivObj = document.createElement("div");
            littleDivObj.style.width = this.widthArr[j] + "px";
            littleDivObj.style.height = this.heightArr[0] + "px";
            littleDivObj.style.marginLeft = ((leftPix += this.widthArr[(j - 1) < 0 ? 0 : (j - 1)]) + j) + "px";
            littleDivObj.className = "div-title-x";
            littleDivObj.id = "div0_" + j;
            if (count == 1) {
                count++;
                this.titleDivOfY.push(littleDivObj);
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
            this.titleDivOfY.push(littleDivObj);
            node.appendChild(littleDivObj);
        }
        let topPix = -this.heightArr[0];
        count = 1;
        for (let i = 0; i <= this.rowNumber; i++) {
            let littleDivObj = document.createElement("div");
            littleDivObj.style.width = this.widthArr[0] + "px";
            littleDivObj.style.height = this.heightArr[i] + "px";
            littleDivObj.style.marginTop = ((topPix += this.heightArr[(i - 1) < 0 ? 0 : (i - 1)]) + i) + "px";
            littleDivObj.className = "div-title-y";
            littleDivObj.id = "div" + i + "_0";
            if (count == 1) {
                count++;
                this.titleDivOfX.push(littleDivObj);
                node.appendChild(littleDivObj);
                continue;
            }
            littleDivObj.innerHTML = yCode++;
            this.titleDivOfX.push(littleDivObj);
            node.appendChild(littleDivObj);
        }
        this.excelTitleRoot.appendChild(node);
    }

    /**
     * @method createLittleDivInput 
     * create input div DOM
     */
    createLittleDivInput() {
        this.inputDiv = {};
        let littleDivObj = document.createElement("input");
        littleDivObj.style.width = this.widthArr[1] + "px";
        littleDivObj.style.height = this.heightArr[1] + "px";
        littleDivObj.style.marginLeft = "21px";
        littleDivObj.style.marginTop = "21px";
        littleDivObj.className = "div-input";
        littleDivObj.id = "inputDiv";
        littleDivObj.setAttribute("disabled", "disabled");
        this.inputDiv = littleDivObj;
        this.excelInputRoot.appendChild(littleDivObj);
    }

    /**
     * @method createLittleDivTarget 
     * create cell DOM
     */
    createLittleDivTarget() {
        let node = document.createDocumentFragment();
        let leftPix = 1;
        let topPix = this.heightArr[0] + 1;
        this.targetDiv = [{}];
        for (let i = 1; i <= this.rowNumber; i++) {
            for (let j = 1; j <= this.columnNumber; j++) {
                let littleDivObj = document.createElement("div");
                littleDivObj.style.width = this.widthArr[j] + "px";
                littleDivObj.style.height = this.heightArr[i] + "px";
                littleDivObj.style.marginLeft = ((leftPix += this.widthArr[(j - 1)]) + j) + "px";
                littleDivObj.style.marginTop = (topPix + i) + "px";
                littleDivObj.className = "div-target";
                littleDivObj.id = "div" + i + "_" + j;
                littleDivObj.innerHTML=this.contentArr[i][j];
 
                this.targetDiv.push(littleDivObj);
                node.appendChild(littleDivObj);


            }
            leftPix = 1;
            topPix += this.heightArr[i];
        }
        this.excelTargetRoot.appendChild(node);
    }
    
    /**
     * @method extraFunction 
     * add extra function according to the input parameter
     */
    extraFunction() {
        let that = this;
        if (this.selectExcel) {
            new SelectExcel(that);
        }
        if (this.resizeExcel) {
            new ResizeExcel(that);
        }
        if (this.insertOrDeleteExcel) {
            new InsertOrDeleteExcel(that);
        }

    }

    /**
     * @method mouseUp 
     * the mouseUp event function 
     * clean the old title style
     */
    mouseUp(event) {
        if ((this.titleDownID != "")) {
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
            this.titleDownID = "";
            this.createLittleDivTarget();
        }
        this.isDown = false;
    }

    /**
     * @method clean 
     * clean all DOM
     */
    clean() {
        let titleNodeXs = this.host.querySelectorAll(".div-title-x");
        for (let titleX of titleNodeXs) {
            titleX.parentNode.removeChild(titleX);
        }
        let titleNodeYs = this.host.querySelectorAll(".div-title-y");
        for (let titleY of titleNodeYs) {
            titleY.parentNode.removeChild(titleY);
        }
        let targetNodes = this.host.querySelectorAll(".div-target");
        for (let targetNode of targetNodes) {
            targetNode.parentNode.removeChild(targetNode);
        }
    }

    /**
     * @method cleanSelect 
     * clean last operation before a new operation 
     */
    cleanSelect() {
        if (this.isSelected) {
            let clearX = this.titleDivOfX[Number(this.downID.split("_")[0]) + 1];//document.getElementById("div" + that.downID.split("_")[0] + "_0");
            let clearY = this.titleDivOfY[Number(this.downID.split("_")[1]) + 1];//document.getElementById("div0_" + that.downID.split("_")[1]);
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
                let pos = (x - 1) * this.columnNumber + y;
                if (pos < 0) {
                    break;
                }
                this.targetDiv[pos].style.backgroundColor = "white";
            }
        }
        if (this.isDblClicked) {// clean input 
            let inputObj = this.inputDiv;
            inputObj.style.zIndex = 0;
            inputObj.style.opacity = 0;
            inputObj.style.margin = "21px 0 0 21px";
            inputObj.setAttribute("disabled", "disabled");
        }
        if (this.selectTitleID != "") { //clean the whole row or column
            if (this.selectTitleID.split("_")[0] == 0) {
                this.selectTitle.style.color = "black";
                this.selectTitle.style.backgroundColor = "rgba(220,220,220,0.8)";
                this.selectTitle.style.borderTop = "solid 1px #A9A9A9";
                for (let i = 1; i <= this.rowNumber; i++) {
                    let pos = (i - 1) * this.columnNumber + Number(this.selectTitleID.split("_")[1]);
                    this.targetDiv[pos].style.backgroundColor = "white";
                }
            } else if (this.selectTitleID.split("_")[1] == 0) {
                this.selectTitle.style.color = "black";
                this.selectTitle.style.backgroundColor = "rgba(220,220,220,0.8)";
                this.selectTitle.style.borderLeft = "solid 1px #A9A9A9";
                for (let i = 1; i <= this.columnNumber; i++) {
                    let pos = (Number(this.selectTitleID.split("_")[0]) - 1) * this.columnNumber + i;
                    this.targetDiv[pos].style.backgroundColor = "white";
                }
            }
        }
    }
}