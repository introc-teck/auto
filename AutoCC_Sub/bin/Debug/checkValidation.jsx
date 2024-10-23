//수정 - allow -> Modify
//경고 - allow -> Warn
//무시 - allow -> Ignore
var allowText = $.getenv("text");
var allowOpacity = $.getenv("opacity");
var allowOverprint = $.getenv("overprint");
var allowSwatches = $.getenv("swatch");
var isFinished = false;
// 별색 체크
if (app.activeDocument.swatches != null)
{
    var isExistSpotcolor = false;
    for(var i = 0; i < app.activeDocument.swatches.length ; i++) {
        if(app.activeDocument.swatches[i].color.typename == "SpotColor") {
            isExistSpotcolor = true;
        }
    }

    if(isExistSpotcolor) {
        if(allowSwatches == "경고") {
                if(Window.confirm("별색이 존재합니다. 계속 진행하시겠습니까?", false, "모아모아")) {
                    allowSwatches = "무시";
                } else {
                    $.setenv("status", "suspend");
                }
        }
        
        if(allowSwatches == "수정") {
            for (var i = (app.activeDocument.swatches.length - 1); i >= 0; i--)
            {
                if (app.activeDocument.swatches[i].color.typename == "SpotColor")
                {
                    app.activeDocument.swatches[i].remove();
                }
            }
        }
    }
}

/*
var docSelected = app.activeDocument.selection;
if (docSelected.length > 0) {
    if ((docSelected.textFrames.length != 0) || (docSelected.legacyTextItems.length != 0))
    {
        if(allowText == "경고") {
            if(Window.confirm("텍스트 항목이 존재합니다. 계속 진행하시겠습니까?", false, "모아모아")) {
                allowText = "무시";
            } else {
                $.setenv("status", "suspend");
            }
        }
    }
}
*/
for (var i = 0; i < app.activeDocument.pageItems.length; i++)
{
    if(!app.activeDocument.pageItems[i].selected || isFinished) {
        continue;
    }
    var page_item = app.activeDocument.pageItems[i];
    if(page_item.typename == "TextFrame" || page_item.typename == "LegacyTextItem") {
        if(allowText == "경고") {
            if(Window.confirm("텍스트 항목이 존재합니다. 계속 진행하시겠습니까?", false, "모아모아")) {
                allowText = "무시";
            } else {
                isFinished = true;
                $.setenv("status", "suspend");
            }
        }
    }

    if(page_item.opacity != 100) {
        if(allowOpacity == "경고") {
            if(Window.confirm("투명도 개체가 존재합니다. 계속 진행하시겠습니까?", false, "모아모아")) {
                allowOpacity = "무시";
            } else {
                isFinished = true;
                $.setenv("status", "suspend");
            }
        }
    
        if(allowOpacity == "수정") {
            page_item.opacity = 100;
        }
    }

    if (page_item.typename == "PathItem")
    {
        if(page_item.fillOverprint == true || page_item.strokeOverprint == true) {
            if(allowOverprint == "경고") {
                if(Window.confirm("중복인쇄 개체가 존재합니다. 계속 진행하시겠습니까?", false, "모아모아")) {
                    allowOverprint = "무시";
                } else {
                    isFinished = true;
                    $.setenv("status", "suspend");
                }
            }
        }
        
        if(allowOverprint == "수정") {
            page_item.fillOverprint  = false;
            page_item.strokeOverprint  = false;
        }
    }
    }

    if (page_item.typename == "Rastertem")
    {
        if(page_item.overprint == true) {
            if(allowOverprint == "경고") {
                if(Window.confirm("중복인쇄 개체가 존재합니다. 계속 진행하시겠습니까?", false, "모아모아")) {
                        allowOverprint = "무시";
                    } else {
                        isFinished = true;
                        $.setenv("status", "suspend");
                    }
                }
                if(allowOverprint == "수정") {
                    page_item.fillOverprint  = false;
                    page_item.strokeOverprint  = false;
                }
            }
        }

$.setenv("status", "complete");