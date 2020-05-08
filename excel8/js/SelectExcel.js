class SelectExcel {
    constructor(that) {
        that.movingID = "";
        that.dblClickID = "";
        that.dblClicked = {};
        this.bindSelectEvent(that);
    }

    /**
     * @method bindSelectEvent
     * @param {object} that  excel object
     * bind event to excel cell and input
     */
    bindSelectEvent(that) {
        let excelTarget = that.excelTargetRoot;
        excelTarget.addEventListener("mousedown", this.mouseDown.bind(that));
        excelTarget.addEventListener("mouseout", this.mouseOut.bind(that));
        excelTarget.addEventListener("mousemove", this.mouseMove.bind(that));
        excelTarget.addEventListener("dblclick", this.showInput.bind(that));



        let excelInput = that.excelInputRoot;
        excelInput.addEventListener("change", this.changeText.bind(that));
        excelInput.addEventListener("keydown", this.enterSubmit.bind(that));

    }
    
    /**
     * @method mouseDown
     * @param {object} event 
     * change the style of selected cell
     */
    mouseDown(event) {
        event.preventDefault();// prevent drag event
        this.cleanSelect();
        if (event.button == 0) {
            this.menuListRoot.style.display="none";
            this.selected = event.target;
            this.downID = event.target.id.substring(3);
            event.target.style.zIndex = 2;
            event.target.style.border = "2px solid green";

            let titleX = Number(this.downID.split("_")[0]);
            let titleY = Number(this.downID.split("_")[1]);
            let titleTargetX = this.titleDivOfX[titleX + 1]; 
            let titleTargetY = this.titleDivOfY[titleY + 1];
            titleTargetX.style.zIndex = 3;
            titleTargetY.style.zIndex = 3;
            titleTargetX.style.color = "green";
            titleTargetY.style.color = "green";
            titleTargetY.style.borderTop = "2px solid green";
            titleTargetX.style.borderLeft = "2px solid green";
            titleTargetX.style.backgroundColor = "rgba(0,0,0,0.2)";
            titleTargetY.style.backgroundColor = "rgba(0,0,0,0.2)";
            this.input_display.value = titleTargetX.innerText + titleTargetY.innerText;
            this.isSelected = true;
            this.isDown = true;
        }
    }

    /**
     * @method mouseMove
     * @param {object} event 
     * multi-select and change cell style
     */
    mouseMove(event) {
        this.movingID = event.target.id.substring(3);
        if (this.isDown) {
            let movingidArr = this.movingID.split('_');
            let downidArr = this.downID.split('_');
            this.leftID = Math.min(movingidArr[0], downidArr[0]);
            this.leftID = this.leftID > 1 ? this.leftID : 1;
            this.topID = Math.min(movingidArr[1], downidArr[1]);
            this.topID = this.topID > 1 ? this.topID : 1;
            this.rightID = Math.max(movingidArr[0], downidArr[0]);
            this.bottomID = Math.max(movingidArr[1], downidArr[1]);
            if ((this.leftID == this.rightID) && (this.topID == this.bottomID)) {
                return false;
            }
            for (let x = this.leftID; x <= this.rightID; x++) {
                for (let y = this.topID; y <= this.bottomID; y++) {
                    let pos = (x - 1) * this.columnNumber + y;
                    if (pos < 0) {
                        break;
                    }
                    this.targetDiv[pos].style.backgroundColor = "rgba(169, 169, 169, 0.5)";
                }
            }
            this.input_display.value = this.leftID + String.fromCharCode(64 + this.topID) + "*" + this.rightID + String.fromCharCode(64 + this.bottomID);
        }
    }

    /**
     * @method mouseOut
     * prevent the mouse out of host
     */
    mouseOut() {
        if (!this.isDown) {
            return false;
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

    }

    /**
     * @method showInput
     * @param {object} event 
     * the dblclick enevt function 
     * show the input div
     */
    showInput(event) {
        this.selected.style.border = "solid 1px #A9A9A9";
        this.selected.style.zIndex = 0;


        this.isDblClicked = true;
        this.dblClicked = event.target;
        this.dblClickID = this.dblClicked.id.substring(3);
        let firstID = Number(this.dblClickID.split("_")[0]);
        let secondID = Number(this.dblClickID.split("_")[1]);
        let leftMargin = 21;
        let topMargin = 21;
        for (let i = 1; i < firstID; i++) {
            topMargin += (this.heightArr[i] + 1);
        }
        for (let j = 1; j < secondID; j++) {
            leftMargin += (this.widthArr[j] + 1);
        }
        let inputObj = this.inputDiv;//document.getElementById("inputDiv");
        inputObj.value = event.target.innerHTML;
        inputObj.style.width = this.widthArr[secondID] + "px";
        inputObj.style.height = this.heightArr[firstID] + "px";
        inputObj.style.zIndex = 3;
        inputObj.style.opacity = 1;
        inputObj.style.marginTop = topMargin + "px";
        inputObj.style.marginLeft = leftMargin + "px";
        inputObj.style.border = "solid 2px green";

        inputObj.removeAttribute("disabled");
        inputObj.focus();
    }

    /**
     * @method changeText
     * the changeText event 
     * change the cell value after change the input value and store the input value into contentArr
     */
    changeText() {
        let numberStr = this.dblClickID;
        let text = this.inputDiv.value;//document.getElementById("inputDiv").value;
        let firstID = Number(numberStr.split("_")[0]);
        let secondID = Number(numberStr.split("_")[1]);
        let pos = (firstID - 1) * this.columnNumber + secondID;
        let targetObj = this.targetDiv[pos];//document.getElementById("div" + numberStr);
        let updated = false;
        targetObj.innerHTML = text;
        this.contentArr[firstID][secondID]=text;
    }

    /**
     * @method enterSubmit
     * @param {object} event 
     * the keydown event function 
     * change input style after the "Enter" key down
     */
    enterSubmit(event) {
        if (event.keyCode == 13) {
            let inputObj = this.inputDiv//document.getElementById("inputDiv");
            inputObj.style.zIndex = 0;
            inputObj.style.opacity = 0;
            inputObj.style.margin = "21px 0 0 21px";
            inputObj.setAttribute("disabled", "disabled");
        }
    }
}