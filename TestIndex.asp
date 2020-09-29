<!DOCTYPE HTML>
<html>
<head>
    <meta charset="utf-8">
    <title>DBCP Test</title>
    <link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/4.5.2/css/bootstrap.min.css">
    <script src="https://cdnjs.cloudflare.com/ajax/libs/popper.js/1.16.0/umd/popper.min.js"></script>
    <script src="https://maxcdn.bootstrapcdn.com/bootstrap/4.5.2/js/bootstrap.min.js"></script>
    <script src="https://ajax.googleapis.com/ajax/libs/jquery/3.5.1/jquery.min.js"></script>
    <script type="text/javascript">
        $(document).ready(function() {
            var replaceInt = /[0-9]/gi;
            var replaceNotInt = /[^0-9]/gi;

            $('#CommodityValue').keyup(function(event) {
                if(!(event.keyCode >= 37 && event.keyCode <= 40)) {
                    var temp = $(this).val();
                    $(this).val(temp.replace(replaceNotInt, ''));
                }
            });

            $('#CommodityCount').keyup(function(event) {
                if(!(event.keyCode >= 37 && event.keyCode <= 40)) {
                    var temp = $(this).val();
                    $(this).val(temp.replace(replaceNotInt, ''));
                }
            });

            $('#Submit').click(function() {
                var CommodityName = $('#CommodityName').val();

                if (CommodityName.length > 20) {
                    alert('물품명이 너무 깁니다.');
                    return false;
                } else if (CommodityName.match(replaceInt)) {
                    alert('물품명에는 숫자를 넣을 수 없습니다.');
                    return false;
                } else {
                    var CommodityData = {
                        CommodityName   : $('#CommodityName').val(),
                        CommodityValue  : $('#CommodityValue').val(),
                        CommodityCount  : $('#CommodityCount').val()
                    };

                    $.ajax({
                        type:   "POST",
                        url:    "TestProc.asp",
                        data:   CommodityData,
                        success: function(result) {
                            console.log(result)
                            if (result == "success") {
                                alert("등록 완료!");
                            } /* else if(data = "overlap") {
                                alert("중복된 정보입니다. ");
                            }*/ else {
                                alert("등록 실패!")
                            }
                        },
                        error:function(request,status,error){
                            alert("code: "      + request.status        + "\n"
                                + "message: "   + request.responseText  + "\n" 
                                + "error: "     + error);
                        }
                    });
                }
            });
        });
    </script>
</head>
<body>
    <div class="container">
        ajax를 사용한 SPA DBCP 서비스<p>
        물품명: <input type="text" id="CommodityName" name="CommodityName"><br>
        물품가액: <input type="text" id="CommodityValue" name="CommodityValue"> 원<br>
        물품재교수: <input type="text" id="CommodityCount" name="CommodityCount"> 개<br>
        </p>
        <button type="button" id="Submit">등록하기</button>
    </div>
</body>
</html>