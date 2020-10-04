<!-- 
@Designed By RH.BanYeoul

    제약 조건
        1. 로그인 관련
            1-1. 아이디 입력은 6글자 이상 / 20글자 이하
                1-1-1. 아이디는 영문(대소문자), 숫자를 제외한 다른 문자는 기입할 수 없음
            1-2. 패스워드 입력은 6글자 이상 / 30글자 이하
                1-2-1. 패스워드는 영문(대소문자), 숫자, 특수문자를 제외한 다른 문자는 기입할 수 없음
        
        2. 아이디 저장 관련
            2-1. 아이디 저장을 체크하고 로그인을 진행할 경우, 쿠키로 아이디 저장 체크여부, 아이디 정보를 저장
            2-2. 사용자가 Index.asp 홈페이지를 열람할 때마다 아이디 저장 체크여부 쿠키를 조회
                2-2-1. TRUE일 경우 아이디 저장 체크 값을 checked로 설정, 쿠키 값을 통해 Input ID value 지정
        
        3. 자동 로그인 관련
            3-1. 자동 로그인을 체크하고 로그인을 진행하여 성공할 경우, 쿠키로 자동 로그인 체크여부를 저장
            3-2. 사용자가 Index.asp 홈페이지를 열람할 때마다 자동 로그인 체크여부 쿠키를 조회
                3-2-1. TRUE일 경우 자동으로 LoginProcCookies.asp 페이지로 이동
            3-3. 로그아웃 버튼을 누를 경우 쿠키가 삭제되고 체크가 풀려야 함

    ISSUE
        Form Submit 이벤트 고장
 -->

<!DOCTYPE HTML>
<html>
<head>
    <meta charset="utf-8">

    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/materialize/1.0.0/css/materialize.min.css">
    <link rel="stylesheet" href="css/Index.css">
    <link rel="stylesheet" href="css/jquery-sakura.css">

    <script src="https://code.jquery.com/jquery-3.5.1.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/materialize/1.0.0/js/materialize.min.js"></script>
    <script src="js/jquery-sakura.js"></script>
    <script type="text/javascript">
        $(document).ready(function() {
            $('body').sakura();

            $('#loginBox').css('top', '15%');
            $('#loginBox').css('transition', 'all 0.8s');

            $('#ds_user_id').focus(function() {
                $('.input-field #idLabel').css('color', '#DC96BA');
                $('.input-field #idLabel').css('font-weight', 'bold');
                $('input[type="text"]').css('box-shadow', '0 1px 0 white');
                $('input[type="text"]').css('border-bottom', '1px solid #FEDAE8');
            });
            $('#ds_user_id').focusout(function() {
                $('.input-field #idLabel').css('color', '#FEDAE8');
                $('.input-field #idLabel').css('font-weight', 'normal');
                $('input[type="text"]').css('box-shadow', 'none');
                $('input[type="text"]').css('border-bottom', '1px solid gray');
            });
            $('#ds_user_password').focus(function() {
                $('.input-field #pwLabel').css('color', '#DC96BA');
                $('input[type="password"]').css('box-shadow', '0 1px 0 white');
                $('input[type="password"]').css('border-bottom', '1px solid #FEDAE8');
            });
            $('#ds_user_password').focusout(function() {
                $('.input-field #pwLabel').css('color', '#FEDAE8');
                $('input[type="password"]').css('box-shadow', 'none');
                $('input[type="password"]').css('border-bottom', '1px solid gray');
            });
        });
    </script>

    <title>Daily Support [계획적인 삶]</title>
</head>
<body>
    <div class="container">
        <div id="loginBox" class="container center-align">
            <div id="Title">
                <a style="font-size: 48px; font-weight: bold;" href="Index.asp">Daily Support</a>
            </div>
            <div id="formBlock">
                <form method="post" id="LoginForm" action="LoginProc.asp">
                    <div class="input-field" style="margin-top: 13%;">
                        <input type="text" class="validate" id="ds_user_id" name="ds_user_id">
                        <label class="active" id="idLabel" for="ds_user_id">ID</label>
                    </div>
                    <div class="input-field">
                        <input type="password" class="validate" id="ds_user_password" name="ds_user_password">
                        <label class="active" id="pwLabel" for="ds_user_password">Password</label>
                    </div>
                    <div class="input-field" id="checkboxBlock">
                        <span>
                            <label>
                                <input type="checkbox" class="filled-in" id="ds_user_is_save" name="ds_user_is_save">
                                <span>아이디 저장</span>
                            </label>
                        </span>
                        <span>
                            <label>
                                <input type="checkbox" class="filled-in" id="ds_user_is_auto" name="ds_user_is_auto">
                                <span>자동 로그인</span>
                            </label>
                        </span>
                    </div>
                    <div class="input-field" style="margin-top: 10%;">
                        <button type="submit" id="submitButton" class="waves-effect waves-purple btn">로그인</button>
                    </div>
                </form>
            </div>
        </div>
    </div>
</body>
</html>