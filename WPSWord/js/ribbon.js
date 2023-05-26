
//这个函数在整个wps加载项中是第一个执行的
function OnAddinLoad(ribbonUI) {
    if (typeof (wps.ribbonUI) != "object") {
        wps.ribbonUI = ribbonUI
    }

    if (typeof (wps.Enum) != "object") { // 如果没有内置枚举值
        wps.Enum = WPS_Enum
    }

    //默认状态
    wps.PluginStorage.setItem("EnableFlag", true) //往PluginStorage中设置一个标记，用于控制两个按钮的置灰
    wps.PluginStorage.setItem("ApiEventFlag", true) //往PluginStorage中设置一个标记，用于控制ApiEvent的按钮label
    return true
}

var WebNotifycount = 0;
function OnAction(control) {
    const eleId = control.Id
    switch (eleId) {
        case "btnShowMsg":
            {
                const doc = wps.WpsApplication().ActiveDocument
                if (!doc) {
                    alert("当前没有打开任何文档")
                    return
                }
                //alert(doc.Name)
                ChatGPTClicked()
            }
            break;

    }
    return true
}

function GetImage(control) {
    const eleId = control.Id
    switch (eleId) {
        case "btnShowMsg":
            return "images/1.svg"
        case "btnShowDialog":
            return "images/2.svg"
        case "btnShowTaskPane":
            return "images/3.svg"
        default:
            ;
    }
    return "images/newFromTemp.svg"
}

function OnGetEnabled(control) {
    const eleId = control.Id
    switch (eleId) {
        case "btnShowMsg":
            {
                let bFlag = wps.PluginStorage.getItem("EnableFlag")
                return bFlag
                break
            }
        case "btnShowDialog":
            {
                let bFlag = wps.PluginStorage.getItem("EnableFlag")
                return bFlag
                break
            }
        case "btnShowTaskPane":
            {
                let bFlag = wps.PluginStorage.getItem("EnableFlag")
                return bFlag
                break
            }
        default:
            break
    }
    return true
}

function OnGetVisible(control) {
    return true
}

function OnGetLabel(control) {
    const eleId = control.Id
    switch (eleId) {
        case "btnIsEnbable":
            {
                let bFlag = wps.PluginStorage.getItem("EnableFlag")
                return bFlag ? "按钮Disable" : "按钮Enable"
                break
            }
        case "btnApiEvent":
            {
                let bFlag = wps.PluginStorage.getItem("ApiEventFlag")
                return bFlag ? "清除新建文件事件" : "注册新建文件事件"
                break
            }
    }
    return ""
}

function ChatGPTClicked() {
    //按钮状态
    //设置
    wps.PluginStorage.setItem("EnableFlag", false)
    //刷新
    wps.ribbonUI.InvalidateControl("btnShowMsg")

    //if (1 == 1) return;

    //获取选中对象
    let selection = wps.WpsApplication().Selection;
    console.log(selection.Text.length);


    //数据验证
    if (selection.Text.length < 2) {
        alert('请选中文字再试');
        SetBtnEnable();
        return;
    }


    //封装对象
    let data =
    {
        prompt: selection.Text,
        url: 'http://127.0.0.1:3002/conversation'
    };
    //let res = "";
    RequestEndPoint(data, selection);



}

function RequestEndPoint(udata, selection) {
    $.ajax({
        url: udata.url,
        type: 'POST',
        contentType: 'application/json',
        data: JSON.stringify({
            "message": udata.prompt,
            "max_tokens": 2000
        }),
        success: function (data) {
           selection.Text = selection.Text + "\n" + data.response;
             //return (data.response);
            //$('#answer').text(data.response);
        },
        error: function (xhr,status,error) {
            alert(`[${status}] ：[${error}]`);
        },
        complete: function ()
        {
            SetBtnEnable();
        }
    });

}

function SetBtnEnable() {
    //按钮状态
    //设置
    wps.PluginStorage.setItem("EnableFlag", true)
    //刷新
    wps.ribbonUI.InvalidateControl("btnShowMsg")
}