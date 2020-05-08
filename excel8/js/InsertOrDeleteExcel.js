class InsertOrDeleteExcel {
    constructor(that) {
        that.rightClickID = "";
        this.bindInsertOrDeleteExcel(that);
    }

    /**
     * @method bindInsertOrDeleteExcel
     * @param {object} that excel object 
     * bind event to excel title 
     */
    bindInsertOrDeleteExcel(that) {
        let excelTitle = that.excelTitleRoot;
        excelTitle.addEventListener("contextmenu", this.rightClick.bind(that));
        that.insertButton.addEventListener("click", this.insertMenu.bind(that));
        that.deleteButton.addEventListener("click", this.deleteMenu.bind(that));
    }

    /**
     * @method rightClick
     * @param {object} event 
     * show the menyList when right click the title 
     */
    rightClick(event) {
        this.rightClickID = event.target.id;
        let x = event.clientX + this.overflowBody.scrollLeft + document.documentElement.scrollLeft;
        let y = event.clientY - this.marginTopPix + this.overflowBody.scrollTop + document.documentElement.scrollTop;
        let menuList = this.menuListRoot;
        menuList.style.margin = y + "px 0px 0px " + x + "px";
        menuList.style.display="block";
    }

    /**
     * @method insertMenu
     * insert row or column when click insert button where you have right clicked
     */
    insertMenu() {
        let firstNum = Number(this.rightClickID.substring(3).split("_")[0]);
        let secondNum = Number(this.rightClickID.substring(3).split("_")[1]);

        if (firstNum == 0) {//insert column
            if (secondNum == 0) {
                return false;
            }
            this.columnNumber++;
            this.widthArr.splice(secondNum, 0, 60);
            let contentArrNew=this.createContentArr(this.rowNumber,this.columnNumber);
            for(let i=0;i<this.rowNumber+1;i++){
                for(let j=0;j<this.columnNumber;j++){
                    if(j<secondNum){
                        contentArrNew[i][j]=this.contentArr[i][j];
                    }else{
                        contentArrNew[i][j+1]=this.contentArr[i][j];
                    }   
                }

            }
            this.contentArr=contentArrNew;
            this.clean();
            this.createLittleDivTitle();
            this.createLittleDivTarget();
        } else {//insert row
            this.rowNumber++;
            this.heightArr.splice(firstNum, 0, 20);
            let contentArrNew=this.createContentArr(this.rowNumber,this.columnNumber);
            for(let i=0;i<this.rowNumber;i++){
                for(let j=0;j<this.columnNumber+1;j++){
                    if(i<firstNum){
                        contentArrNew[i][j]=this.contentArr[i][j];
                    }else{
                        contentArrNew[i+1][j]=this.contentArr[i][j];
                    }
                    
                }
            }

            this.contentArr=contentArrNew;
            this.clean();
            this.createLittleDivTitle();
            this.createLittleDivTarget();
        }
        this.menuListRoot.style.display="none";
    }

    /**
     * @method deleteMenu
     * delete clicked row or column when click delete button
     */
    deleteMenu() {
        let firstNum = Number(this.rightClickID.substring(3).split("_")[0]);
        let secondNum = Number(this.rightClickID.substring(3).split("_")[1]);
        if (firstNum == 0) {//delete column
            if (secondNum == 0) {
                return false;
            }
            this.columnNumber--;
            this.widthArr.splice(secondNum, 1);
            let contentArrNew=this.createContentArr(this.rowNumber,this.columnNumber);

            for(let i=0;i<this.rowNumber+1;i++){
                for(let j=0;j<this.columnNumber;j++){
                    if(j<secondNum){
                        contentArrNew[i][j]=this.contentArr[i][j];
                    }else{
                        contentArrNew[i][j]=this.contentArr[i][j+1];
                    }
                }
            }
            this.contentArr=contentArrNew;
            this.clean();
            this.createLittleDivTitle();
            this.createLittleDivTarget();
        } else {//delete row
            this.rowNumber--;
            this.heightArr.splice(firstNum, 1);

            let contentArrNew=this.createContentArr(this.rowNumber,this.columnNumber);

            for(let i=0;i<this.rowNumber+1;i++){
                for(let j=0;j<this.columnNumber;j++){
                    if(i<firstNum){
                        contentArrNew[i][j]=this.contentArr[i][j];
                    }else{
                        contentArrNew[i][j]=this.contentArr[i+1][j];
                    }
                }
            }
            this.contentArr=contentArrNew;
            this.clean();
            this.createLittleDivTitle();
            this.createLittleDivTarget();
        }
        this.menuListRoot.style.display="none";
    }
}