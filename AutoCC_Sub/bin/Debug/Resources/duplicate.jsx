//app.preferences.setBooleanPreference('aiFileFormat/fixBadPluginGlobalObjectNames', true);

var side = Number($.getenv("side"));
var num = Number($.getenv("num"));
var width = parseFloat($.getenv("width")) / 0.352777778;
var height = parseFloat($.getenv("height")) / 0.352777778;
var tolerance = 0.5 / 0.352777778;

/*
side = 2;
num = 1;
width = 185 / 0.352777778;
height = 260 / 0.352777778;
tolerance = 0.5 / 0.352777778
*/

var newItem;
var docSelected = app.activeDocument.selection;
if(side * num != docSelected.length) {
    alert('선택된 그룹의 갯수가 일치하지 않습니다');
} else {

if (docSelected.length > 0) {
    // Create a new document and move the selected items to it.
    var docPreset = new DocumentPreset;
     docPreset.width =width;
     docPreset.height = height;
     docPreset.units = RulerUnits.Millimeters;
     docPreset.previewMode = DocumentPreviewMode.PixelPreview;
     docPreset.rasterResolution = DocumentRasterResolution.MediumResolution;
     docPreset.presetType = DocumentPresetType.Web;
     docPreset.artboardRowsOrCols = side;
     docPreset.numArtboards = side * num;
    var newDoc = app.documents.addDocument(DocumentColorSpace.CMYK, docPreset);
    newDoc.activate();

    if (docSelected.length > 0) {
        docSelected = organize(docSelected);
        for (var i = 0; i < docSelected.length; i++) {
            docSelected[i].selected = false;
            newItem = docSelected[i].duplicate(newDoc, ElementPlacement.PLACEATBEGINNING);
            var bound = getRealVisibleBounds(newItem);
            //var bound = getVisibleBounds(newItem);
            if(bound == null) bound = [0,0,0,0];
            var page_left = bound[0];
            var page_top = bound[1];
            var page_width = (bound[2] - bound[0]);
            var page_height = (bound[1] - bound[3]);
            var gap_left = 0;
            var gap_top = 0;
            
            if ((page_width > (width - tolerance)) && (page_width < (width + tolerance)) && (page_height > (height - tolerance)) && (page_height < (height + tolerance))) {
                gap_left = newItem.left - bound[0];
                gap_top = newItem.top - bound[1];
            } else {
                gap_left = -((newItem.width - width) / 2);
                gap_top = -((newItem.height - height) / 2);
                //alert(gap_left + " " + gap_top);
                for(var j = 0; j < newItem.pageItems.length; j++) {
                    var page_left1 = newItem.pageItems[j].left;
                    var page_top1 = newItem.pageItems[j].top;
                    var page_width1 = newItem.pageItems[j].width;
                    var page_height1 = newItem.pageItems[j].height;
                    if ((page_width1 > (width - tolerance)) && (page_width1 < (width + tolerance)) && (page_height1 > (height - tolerance)) && (page_height1 < (height + tolerance))) {
                        if(gap_left < newItem.left - page_left1) {
                            gap_left = newItem.left - page_left1;
                        }
                        if(gap_top < newItem.top - page_top1) {
                            gap_top = newItem.top + page_top1;
                        }
                    }
                }
            }

            newItem.left = newDoc.artboards[i].artboardRect[0] + gap_left;
            newItem.top = newDoc.artboards[i].artboardRect[1] + gap_top;
        }
    
    } else {
        docSelected.selected = false;
        newItem = docSelected.parent.duplicate(newDoc, ElementPlacement.PLACEATBEGINNING);
    }
} else {
//alert("Please select one or more art objects");
}
}

function getRealVisibleBounds(grp) {
     var outerBounds = [];
     for(var i = grp.pageItems.length - 1; i >= 0;i--)  {
          var bounds = [];
          
          if ((grp.pageItems[i].width > (width - tolerance)) && (grp.pageItems[i].width < (width + tolerance)) && (grp.pageItems[i].height > (height - tolerance)) && (grp.pageItems[i].height < (height + tolerance))) {
                return grp.pageItems[i].visibleBounds;
          }
          else if(grp.pageItems[i].typename == 'GroupItem') {
               bounds =  getRealVisibleBounds(grp.pageItems[i]);
          }
          else if((grp.pageItems[i].typename == 'PathItem' || grp.pageItems[i].typename == 'CompoundPathItem') 
               && (grp.pageItems[i].clipping || !grp.clipped)) {
               bounds = grp.pageItems[i].visibleBounds;
          }
          if (bounds != null) {
               outerBounds = maxBounds(outerBounds,bounds);
          }
     }
     return (outerBounds.length == 0) ? null : outerBounds;
}

function maxBounds(ary1,ary2) {
     var res = [];
     if(ary1.length == 0)
          res = ary2;
     else if(ary2.length == 0)
          res = ary1;
     else {
          res[0] = Math.min(ary1[0],ary2[0]);
          res[1] = Math.max(ary1[1],ary2[1]);
          res[2] = Math.max(ary1[2],ary2[2]);
          res[3] = Math.min(ary1[3],ary2[3]);
     }
     return res;
}

function positionVisible(grp,x,y)
{
     var bounds = getRealVisibleBounds(grp);
     var newX = x + (grp.left - bounds[0]);
     var newY = y + (grp.top - bounds[1]);
     grp.position = [newX,newY];
}


function organize(selection) {
    var groups = selection;
    var currentRowMarker;
    var groupList = []; //array of all groupItems
    var sortedGroupList = []; //array of subarrays sorted by visible bounds
    var temp = []; //temporary array for the current row of groupItems. sort this array from left to right first, then push the entire array into "sortedGroupList"

    for (var g=0; g < groups.length; g++){
        sortedGroupList.push(groups[g]);
    }

    sortedGroupList.sort(function(a, b) { // 오름차순
        //return (a.top * 100 + a.left) - (b.top * 100 + b.left);
        return (b.top * 100 - b.left) - (a.top * 100 - a.left);
    });

    return sortedGroupList;

}