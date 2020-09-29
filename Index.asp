<!DOCTYPE HTML>
<html>
<head>
    <meta charset="utf-8">

    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/materialize/1.0.0/css/materialize.min.css">
    <link rel="stylesheet" href="css/Index.css">

    <script src="https://code.jquery.com/jquery-3.5.1.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/materialize/1.0.0/js/materialize.min.js"></script>
    <script type="text/javascript">
        $(document).ready(function() {
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
                <a style="font-size: 36px; font-weight: bold;">Daily Support</a>
            </div>
            <div id="formBlock">
                <form method="post" action="#">
                    <div class="input-field">
                        <input type="text" class="validate" id="ds_user_id" name="ds_user_id">
                        <label class="active" id="idLabel" for="ds_user_id">ID</label>
                    </div>
                    <div class="input-field">
                        <input type="password" class="validate" id="ds_user_password" name="ds_user_password">
                        <label class="active" id="pwLabel" for="ds_user_password">Password</label>
                    </div>
                    <div class="input-field" style="margin-top: 10%;">
                        <input type="checkbox" class="filled-in" id="ds_user_is_save" name="ds_user_is_save">
                        <span>아이디 저장</span>
                        <input type="checkbox" class="filled-in" id="ds_user_is_auto" name="ds_user_is_auto">
                        <span>자동 로그인</span>
                    </div>
                    <div class="input-field" style="margin-top: 10%;">
                        <button type="submit" id="submitButton" class="waves-effect waves-light btn">로그인</button>
                    </div>
                </form>
            </div>
        </div>
    </div>
</body>
</html>