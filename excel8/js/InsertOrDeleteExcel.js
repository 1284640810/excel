class InsertOrDeleteExcel{
    constructor(that){
        that.rightClickID = "";
        this.bindInsertOrDeleteExcel(that);
    }
    bindInsertOrDeleteExcel(that){
        let excelTitle = that.excelTitleRoot;
        excelTitle.addEventListener("contextmenu",this.rightClick.bind(that));
        that.insertButton.addEventListener("click",this.insertMenu.bind(that));
        that.deleteButton.addEventListener("click", this.deleteMenu.bind(that));
    }
    rightClick(event) {
        this.rightClickID = event.target.id;
        let x = event.clientX+this.overflowBody.scrollLeft+document.documentElement.scrollLeft;
        let y = event.clientY - this.marginTopPix+this.overflowBody.scrollTop+document.documentElement.scrollTop;
        let menuList = this.menuListRoot;
        menuList.style.margin = y + "px 0px 0px " + x + "px";
        menuList.style.opacity = 1;
        menuList.removeAttribute("disabled");
    }

    insertMenu(event){
        let firstNum = Number(this.rightClickID.substring(3).split("_")[0]);
        let secondNum = Number(this.rightClickID.substring(3).split("_")[1]);
        let contentArr;
        if (firstNum == 0) {//添加列
            if(secondNum==0){
                return false;
            }
            this.columnNumber++;
            this.widthArr.splice(secondNum, 0, 60);
            for (let i = 0; i < this.contentList.length; i++) {
                contentArr = this.contentList[i].split("_");
                if (secondNum <= contentArr[1]) {
                    this.contentList[i] = contentArr[0] + "_" + (Number(contentArr[1]) + 1) + "_" + contentArr[2];
                }
            }
            this.clean();
            //InsertOrDeleteExcel.clean();
            this.createLittleDivTitle();
            this.createLittleDivTarget();
        } else {//添加行
            this.rowNumber++;
            this.heightArr.splice(firstNum, 0, 20);
            for (let i = 0; i < this.contentList.length; i++) {
                contentArr = this.contentList[i].split("_");
                if (firstNum <= contentArr[0]) {
                    this.contentList[i] = (Number(contentArr[0]) + 1) + "_" + contentArr[1] + "_" + contentArr[2];
                }
            }
            this.clean();
            //InsertOrDeleteExcel.clean();
            this.createLittleDivTitle();
            this.createLittleDivTarget();
        }
        this.menuListRoot.style.opacity = 0;
        this.menuListRoot.style.margin = "-400px";
        this.menuListRoot.setAttribute("disabled", "disabled");
    }
    deleteMenu(event){
        let firstNum = Number(this.rightClickID.substring(3).split("_")[0]);
        let secondNum = Number(this.rightClickID.substring(3).split("_")[1]);
        let contentArr;
        if (firstNum == 0) {//删除列
            if(secondNum==0){
                return false;
            }
            this.columnNumber--;
            this.widthArr.splice(secondNum, 1);
            for (let i = 0; i < this.contentList.length; i++) {
                contentArr = this.contentList[i].split("_");
                if (secondNum < contentArr[1]) {
                    this.contentList[i] = contentArr[0] + "_" + (Number(contentArr[1]) - 1) + "_" + contentArr[2];
                } else if (secondNum == contentArr[1]) {
                    this.contentList.splice(i, 1);
                    i--;
                }
            }
            this.clean();
            //InsertOrDeleteExcel.clean();
            this.createLittleDivTitle();
            this.createLittleDivTarget();
        } else {//删除行
            this.rowNumber--;
            this.heightArr.splice(firstNum, 1);
            for (let i = 0; i < this.contentList.length; i++) {
                contentArr = this.contentList[i].split("_");
                if (firstNum < contentArr[0]) {
                    this.contentList[i] = (Number(contentArr[0]) - 1) + "_" + contentArr[1] + "_" + contentArr[2];
                } else if (firstNum == contentArr[0]) {
                    this.contentList.splice(i, 1);
                    i--;
                }
            }
            this.clean();
            //InsertOrDeleteExcel.clean();
            this.createLittleDivTitle();
            this.createLittleDivTarget();
        }
        this.menuListRoot.style.opacity = 0;
        this.menuListRoot.style.margin = "-400px";
        this.menuListRoot.setAttribute("disabled", "disabled");
    }
}