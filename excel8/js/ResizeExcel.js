class ResizeExcel {
    constructor(that) {
        that.direction = "";
        that.disW = 0;
        that.disL = 0;
        that.disH = 0;
        that.disT = 0;
        that.disX = 0;
        that.disY = 0;
        this.bindResizeEvent(that);
    }

    /**
     * @method bindResizeEvent
     * @param {object} that 
     * bind event to excel title 
     */
    bindResizeEvent(that) {
        let excelTitle = that.excelTitleRoot;
        excelTitle.addEventListener("click", this.clickTitle.bind(that));
        excelTitle.addEventListener("mousedown", this.titleMouseDown.bind(that));
        excelTitle.addEventListener("mousemove", this.titleDownMove.bind(that));
        excelTitle.addEventListener("mouseout", function (event) {
            event.target.style.backgroundColor = "rgba(220,220,220,0.8)";
        });

    }

    /**
     * @method clickTitle
     * @param {object} event 
     * select the whole row or column when click the title
     */
    clickTitle(event) {
        //console.log("click title");
        event.preventDefault();
        let titleID = event.target.id.substring(3);
        let titleTarget = event.target;
        this.cleanSelect();
        if (titleID.split("_")[0] == 0) {//选择整列
            if (titleID.split("_")[1] == 0) {
                return false;
            }
            titleTarget.style.color = "green";
            titleTarget.style.backgroundColor = "rgba(0,0,0,0.2)";
            titleTarget.style.borderTop = "2px solid green";
            for (let i = 1; i <= this.rowNumber; i++) {
                let pos = (i - 1) * this.columnNumber + Number(titleID.split("_")[1]);
                this.targetDiv[pos].style.backgroundColor = "rgba(169, 169, 169, 0.5)";
            }
            this.input_display.value = "1" + titleTarget.innerText;
        } else if (titleID.split("_")[1] == 0) {
            titleTarget.style.color = "green";
            titleTarget.style.backgroundColor = "rgba(0,0,0,0.2)";
            titleTarget.style.borderLeft = "2px solid green";
            for (let i = 1; i <= this.columnNumber; i++) {
                let pos = (Number(titleID.split("_")[0]) - 1) * this.columnNumber + i;
                //console.log("click y"+pos);
                this.targetDiv[pos].style.backgroundColor = "rgba(169, 169, 169, 0.5)";
            }
            this.input_display.value = titleTarget.innerText + "A";
        }
        this.selectTitle = titleTarget;
        this.selectTitleID = titleID;
    }

    /**
     * @method titleMouseDown
     * @param {object} event 
     * begin resize 
     */
    titleMouseDown(event) {
        event.preventDefault();
        if (event.button != 0) {
            return false;
        }
        this.titleDowned = event.target;
        this.titleDownID = event.target.id;
        this.disW = this.titleDowned.offsetWidth;
        this.disL = this.titleDowned.offsetLeft;
        this.disH = this.titleDowned.offsetHeight;
        this.disT = this.titleDowned.offsetTop;
        this.disX = event.clientX + this.overflowBody.scrollLeft + document.documentElement.scrollLeft;
        this.disY = event.clientY - this.marginTopPix + this.overflowBody.scrollTop + document.documentElement.scrollTop;
        if (event.target.className == "div-title-y") {//vertical resize 
            if (this.disY > this.disT + this.disH - 5) {
                this.direction = "down";

            } else if ((this.disY > this.disT) && (this.disY < this.disT + 5)) {
                this.direction = "down2"
                let num = Number(this.titleDownID.substring(3).split("_")[0]) - 1;
                this.titleDownID = "div" + num + "_0";
                this.titleDowned = this.titleDivOfX[num + 1];
                this.disT = this.titleDowned.offsetTop;
                this.disH = this.titleDowned.offsetHeight;
            } else {
                this.titleDownID = "";
            }
        } else if (event.target.className == "div-title-x") {//horizontal resize
            if (this.disX > this.disL + this.disW - 8) {
                this.direction = "right";
            } else if ((this.disX < this.disL + 8) && (this.disX > this.disL)) {
                this.direction = "right2";
                let num = Number(this.titleDownID.substring(3).split("_")[1]) - 1;
                this.titleDownID = "div0_" + num;
                this.titleDowned = this.titleDivOfY[num + 1];
                this.disW = this.titleDowned.offsetWidth;
                this.disL = this.titleDowned.offsetLeft;
            } else {
                this.titleDownID = "";
            }
        }
    }

    /**
     * @method titleDownMove
     * @param {object} event 
     * resizing
     */
    titleDownMove(event) {

        if (this.titleDownID != "") {
            switch (this.direction) {
                case "right2": {
                    this.titleDowned.style.width = this.disW + (event.clientX + this.overflowBody.scrollLeft + document.documentElement.scrollLeft - this.disX) + 'px';
                    break;
                }
                case "right": {
                    this.titleDowned.style.width = this.disW + (event.clientX + this.overflowBody.scrollLeft + document.documentElement.scrollLeft - this.disX) + 'px';  
                    break;
                }
                case "down": {
                    this.titleDowned.style.height = this.disH + (event.clientY - this.marginTopPix + this.overflowBody.scrollTop + document.documentElement.scrollTop - this.disY) + "px";
                    break;
                }
                case "down2": {
                    this.titleDowned.style.height = this.disH + (event.clientY - this.marginTopPix + this.overflowBody.scrollTop + document.documentElement.scrollTop - this.disY) + "px";
                    break;
                }
            }
        } else {// change cursor
            event.preventDefault();
            event.target.style.backgroundColor = "rgba(144,238,144,0.8)";
            let titleMoved = event.target;
            let disW = titleMoved.offsetWidth;
            let disX = event.clientX + this.overflowBody.scrollLeft + document.documentElement.scrollLeft;

            let disL = titleMoved.offsetLeft;
            let disH = titleMoved.offsetHeight;
            let disT = titleMoved.offsetTop;
            let disY = event.clientY - this.marginTopPix + this.overflowBody.scrollTop + document.documentElement.scrollTop;
            if (event.target.className == "div-title-y") {
                if ((disY > (disT + disH - 5)) || ((disY > disT) && (disY < disT + 5))) {
                    titleMoved.style.cursor = "s-resize";

                } else {
                    titleMoved.style.cursor = "pointer";

                }
            } else if (event.target.className == "div-title-x") {
                if ((disX > disL + disW - 8) || ((disX < disL + 8) && (disX > disL))) {
                    titleMoved.style.cursor = "w-resize";

                } else {
                    titleMoved.style.cursor = "pointer";

                }
            }
        }

    }
}