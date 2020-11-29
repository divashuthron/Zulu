//Ajax
$.Ajax4Form = function (form, obj, debuge) {
    debuge = debuge || false;
    var clearForm = obj["clear"] || false;
    var resetForm = obj["reset"] || false;
    var viewLoding = obj["Loding"] || true;
    
    //-------------- alert ---------------
    if (debuge) {
        alert(" url: " + obj["url"]
        + "\r\n param :" + obj["param"]
        + "\r\n dataType :" + obj["dataType"]
        + "\r\n before :" + obj["before"]
        + "\r\n success :" + obj["success"]
        + "\r\n complete :" + obj["complete"]
        + "\r\n clear :" + obj["clear"]
        + "\r\n reset :" + obj["reset"]);
    }
    //-------------- alert ---------------
    $(form).ajaxForm({
        url: obj["url"] + obj["param"], // override for form's 'action' attribute
        type: "post", 					// 'get' or 'post', override for form's 'method' attribute
        dataType: obj["dataType"], 		// 'xml', 'script', or 'json' (expected server response type)
        clearForm: clearForm, 			// clear all form fields after successful submit
        resetForm: resetForm, 			// reset the form after successful submit
        beforeSubmit: function () {
            //eval(obj["before"]);
        },
        success: function (datas, state) {
            eval(obj["success"]);
        },
        error: function (reason, e) {
            alert('서버연결에 실패했습니다. : ' + e);
        },
        complete: function () {
            //eval(obj["complete"]);
        }
    });
}