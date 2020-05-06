class ResizeExcel {
    constructor(that){
        that.direction="";
        that.disW=0;
        that.disL=0;
        that.disH=0;
        that.disT=0;
        that.disX=0;
        that.disY=0;
        this.bindResizeEvent(that);
    }
    bindResizeEvent(that) {
        let excelTitle = that.excelTitleRoot;
        excelTitle.addEventListener("click",this.clickTitle.bind(that));
        excelTitle.addEventListener("mousedown",this.titleMouseDown.bind(that));
        excelTitle.addEventListener("mousemove",this.titleDownMove.bind(that));
        excelTitle.addEventListener("mouseout", function (event) {
            event.target.style.backgroundColor = "rgba(220,220,220,0.8)";
        });

    }
    clickTitle(event) {
        //console.log("click title");
        event.preventDefault();
        let titleID = event.target.id.substring(3);
        let titleTarget = event.target;
        this.cleanSelect();
        if (titleID.split("_")[0] == 0) {//选择整列
            if(titleID.split("_")[1] == 0){
                return false;
            }
            titleTarget.style.color = "green";
            titleTarget.style.backgroundColor = "rgba(0,0,0,0.2)";
            titleTarget.style.borderTop = "2px solid green";
            for (let i = 1; i <= this.rowNumber; i++) {
                let pos=(i-1)*this.columnNumber+Number(titleID.split("_")[1]);
                //console.log("click x"+pos);
                this.targetDiv[pos].style.backgroundColor = "rgba(169, 169, 169, 0.5)";//document.getElementById("div" + i + "_" + titleID.split("_")[1]).style.backgroundColor = "rgba(169, 169, 169, 0.5)";
            }
            this.input_display.value = "1" + titleTarget.innerText;
        } else if(titleID.split("_")[1] == 0){
            titleTarget.style.color = "green";
            titleTarget.style.backgroundColor = "rgba(0,0,0,0.2)";
            titleTarget.style.borderLeft = "2px solid green";
            for (let i = 1; i <= this.columnNumber; i++) {
                let pos=(Number(titleID.split("_")[0])-1)*this.columnNumber+i;
                //console.log("click y"+pos);
                this.targetDiv[pos].style.backgroundColor = "rgba(169, 169, 169, 0.5)";//document.getElementById("div" + titleID.split("_")[0] + "_" + i).style.backgroundColor = "rgba(169, 169, 169, 0.5)";
            }
            this.input_display.value = titleTarget.innerText + "A";
        }
        this.selectTitle = titleTarget;
        this.selectTitleID = titleID;
    }

    titleMouseDown(event) {
        event.preventDefault();
        if (event.button != 0) {
            return false;
        }
        this.titleDowned = event.target;
        this.titleDownID = event.target.id;
        this.disW = this.titleDowned.offsetWidth;//当前titlediv宽度
        this.disL = this.titleDowned.offsetLeft;//当前titlediv左边界距屏幕左侧距离
        this.disH = this.titleDowned.offsetHeight;
        this.disT = this.titleDowned.offsetTop;
        this.disX = event.clientX+this.overflowBody.scrollLeft+document.documentElement.scrollLeft;//当前鼠标x坐标
        this.disY = event.clientY - this.marginTopPix+this.overflowBody.scrollTop+document.documentElement.scrollTop;//当前鼠标y坐标
        if (event.target.className == "littleDivTitleY") {//垂直拉伸
            if (this.disY > this.disT + this.disH - 5) {
                this.direction = "down";

            } else if ((this.disY > this.disT) && (this.disY < this.disT + 5)) {
                this.direction = "down2"
                let num=Number(this.titleDownID.substring(3).split("_")[0])-1;
                // console.log("title y id"+num);
                this.titleDownID = "div" +num+"_0";
                this.titleDowned = this.titleDivOfX[num+1]//document.getElementById(this.titleDownID);
                this.disT = this.titleDowned.offsetTop;
                this.disH = this.titleDowned.offsetHeight;
            }else {
                this.titleDownID="";
            }
        } else if (event.target.className == "littleDivTitleX") {//水平拉伸
            if (this.disX > this.disL + this.disW - 8) {
                this.direction = "right";
            } else if ((this.disX < this.disL + 8) && (this.disX > this.disL)) {
                this.direction = "right2";
                let num=Number(this.titleDownID.substring(3).split("_")[1])-1;
                // console.log("title x id"+num);
                this.titleDownID = "div0_" +num;
                this.titleDowned = this.titleDivOfY[num+1]//document.getElementById(this.titleDownID);
                this.disW = this.titleDowned.offsetWidth;//当前titlediv宽度
                this.disL = this.titleDowned.offsetLeft;//当前titlediv左边界距屏幕左侧距离
            }else {
                this.titleDownID="";
            }
        }
    }

    titleDownMove(event) {
        
        if(this.titleDownID!=""){
            // document.style.cursor="pointer";
            switch (this.direction) {
                case "right2": {
                    this.titleDowned.style.width = this.disW + (event.clientX+this.overflowBody.scrollLeft +document.documentElement.scrollLeft- this.disX) + 'px';
                    break;
                }
                case "right": {
                    this.titleDowned.style.width = this.disW + (event.clientX+this.overflowBody.scrollLeft +document.documentElement.scrollLeft- this.disX) + 'px';  //右侧元素的宽度=原来的宽度+改变的宽度
                    break;
                }
                case "down": {
                    this.titleDowned.style.height = this.disH + (event.clientY - this.marginTopPix + this.overflowBody.scrollTop+document.documentElement.scrollTop - this.disY) + "px";
                    break;
                }
                case "down2": {
                    this.titleDowned.style.height =this.disH + (event.clientY - this.marginTopPix + this.overflowBody.scrollTop+document.documentElement.scrollTop - this.disY) + "px";
                    break;
                }
            }
        }else{
            event.preventDefault();
            event.target.style.backgroundColor="rgba(144,238,144,0.8)";
            let titleMoved = event.target;
            let disW = titleMoved.offsetWidth;//当前titlediv宽度
            let disX = event.clientX+this.overflowBody.scrollLeft+document.documentElement.scrollLeft;//当前鼠标x坐标
            
            let disL = titleMoved.offsetLeft;//当前titlediv左边界距屏幕左侧距离
            let disH = titleMoved.offsetHeight;
            let disT = titleMoved.offsetTop;
            let disY = event.clientY - this.marginTopPix+this.overflowBody.scrollTop+document.documentElement.scrollTop;//当前鼠标y坐标

            // console.log(that.overflowBody.scrollLeft)
            // console.log(that.overflowBody.scrollTop);
            // //let disX=event.pageX;
            // console.log(that.disX);
            // console.log(that.disY);
            if (event.target.className == "littleDivTitleY") {
                if ((disY > (disT + disH - 5)) || ((disY > disT) && (disY < disT + 5))) {
                    titleMoved.style.cursor = "s-resize";
                    
                } else {
                    titleMoved.style.cursor = "pointer";
                    
                }
            } else if(event.target.className == "littleDivTitleX"){
                if ((disX > disL + disW - 8) || ((disX < disL + 8) && (disX > disL))) {
                    titleMoved.style.cursor = "w-resize";
                    
                } else {
                    titleMoved.style.cursor = "pointer";
                    
                }
            }
        }

    }
}