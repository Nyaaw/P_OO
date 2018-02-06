



// please add your custiomize functions to this method.
var initCustomize = function () {
  
    addDemo();

    addRestfulAPI();

    enableDimensionTextRefresh();
};

function addDemo() {
    var myIcon = new CToolbarIcon();
    myIcon.id = "myIcon";
    myIcon.event = "saveOnServer()";
    myIcon.title = "Save on server";
    myIcon.reqDoc = true;


    var myGroup = new CToolbarIconGroup();
    myGroup.name = "Save Server";
    myGroup.addIcon(myIcon);

    var myTab = new CToolbar();
    myTab.id = "Customizebar";
    myTab.name = "Customize";
    myTab.addIconGroup(myGroup);


    addToolbarTab(myTab);
}

function addRestfulAPI() {
    var myIcon = new CToolbarIcon();
    myIcon.id = "SaveRestful";
    myIcon.event = "saveOnRestfulBack()";
    myIcon.title = "Save and call back.";
    myIcon.reqDoc = true;

    var myIcon2 = new CToolbarIcon();
    myIcon2.id = "ExportRestful";
    myIcon2.event = "showExportPanel()";
    myIcon2.title = "Export and call back.";
    myIcon2.reqDoc = true;

    var myGroup = new CToolbarIconGroup();
    myGroup.name = "Save";
    myGroup.addIcon(myIcon);
    myGroup.addIcon(myIcon2);

    var myTab = new CToolbar();
    myTab.id = "RestFulAPI";
    myTab.name = "RestFul API";
    myTab.addIconGroup(myGroup);







    addToolbarTab(myTab);
}

// demo function to save the file on server
function saveOnServer() {
    // whether the file is loaded
    if (getCurrentFileId() == "")
        return;
    // get the basic save datas from API(getSaveDatas)
    var datas = getSaveDatas();
    // set the customize calling function in userCommandProcessHandler
    datas.action = "SaveFileOnServer";
    var options = {
        type: "POST",
        url: getServerHandlerUrl(),
        async: false,
        data: datas,
        success: function (result) {
            // hide the loading panel
            $("#loading").hide();
            // get the message
            var arr = eval('(' + result + ')');
            if (arr.state == "success") {
                saveFile = arr.msg;
                // open the file that store on server
                window.open(getCacheFileFolder(getCurrentFileId()) + saveFile);
            }
            else {
                alert(arr.msg);
            }
        },
        error: function (err) {
            $("#loading").hide();
        }
    }
    // show the loading panel
    $("#loading").show();
    // send the message
    $.ajax(options);
}
var getAreaLocation = function (location) {
    //alert(location.pageIndex+":"+location.x + "," + location.y + "  " + location.w + ":" + location.h);
    if (window.parent!=null&&window.parent.ocrarea != null)
        window.parent.ocrarea(location);
}

function saveOnRestfulBack() {
    // whether the file is loaded
    if (getCurrentFileId() == "")
        return;
    // get the basic save datas from API(getSaveDatas)
    var datas = getSaveDatas();
    datas.action = "SaveFile";
    var options = {
        type: "POST",
        url: getServerHandlerUrl(),
        async: false,
        data: datas,
        success: function (result) {
            // hide the loading panel
            $("#loading").hide();
            // get the message
            var arr = eval('(' + result + ')');
            if (arr.state == "success") {
                saveFile = arr.msg;
                // open the file that store on server
                alert("Save on server ok.");
            }
            else {
                alert(arr.msg);
            }
        },
        error: function (err) {
            $("#loading").hide();
        }
    }
    // show the loading panel
    $("#loading").show();
    // send the message
    $.ajax(options);
}
var isshow = false;
function showExportPanel() {
    isshow = !isshow;
    if (isshow) {
        var str = "<div id='_restfuloptionsPanel' class='popboxbase' style='Margin:100px 140px 0px 0px;border:1px solid #A0A0A0;box-shadow:2px 2px 2px #666;z-index:600;'></div>";
        $("body").append(str);
        $("#_optionsPanel").css({
            width: "140px",
            height: "120px",
            display: "block"
        });
        var table = "<table style='color:#444444;width:100%;height:100%;cursor:pointer;'>";
        table += "<tr><td id='restful_docx' ><a href='#' onclick='exportOnRestFul(\"docx\");return false;' >Microsoft Word ...</a></td></tr>";
        table += "<tr><td id='restful_pdf'><a href='#' onclick='exportOnRestFul(\"pdf\");return false;'>PDF ...</a></td></tr>";
        table += "<tr><td id='restful_tiff' ><a href='#' onclick='exportOnRestFul(\"tiff\");return false;'>Tiff ...</a></td></tr></table>";
        $("#_restfuloptionsPanel").append(table);
        $("#_optionsPanel").hide();
    } else {
        $("#_restfuloptionsPanel").remove();
    }
    
}
function exportOnRestFul(operation) {
    // whether the file is loaded
    if (getCurrentFileId() == "")
        return;
    // get the basic save datas from API(getSaveDatas)
    var datas = getSaveDatas();
    if (operation == "tiff") {
        datas.type = "Tiff";
        var imgOptions = new ImageOutOptions();
        imgOptions.colorSpace = "Determine Automatically";
        imgOptions.compressMode = "LZW";
        imgOptions.resolution = 96;
        datas.imgOptions = $.toJSON(imgOptions);
    }
    else if (operation == "docx")
        datas.type = "Word";
    else if (operation == "pdf")  
        datas.type = "pdf";
        else return;
    datas.action = "ExportTo";
    $("#_restfuloptionsPanel").remove();
    var options = {
        type: "POST",
        url: getServerHandlerUrl(),
        async: false,
        data: datas,
        success: function (result) {
            // hide the loading panel
            $("#loading").hide();
            // get the message
            var arr = eval('(' + result + ')');
            if (arr.state == "success") {
                saveFile = arr.msg;
                // open the file that store on server
                alert("Export on server ok.");
            }
            else {
                alert(arr.msg);
            }
        },
        error: function (err) {
            $("#loading").hide();
        }
    }
    // show the loading panel
    $("#loading").show();
    // send the message
    $.ajax(options);
    isshow = false;
}