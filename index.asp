<!-- #Include Virtual = "/Include/Header.asp" -->
	<div id="colorlib-page">
		<a href="#" class="js-colorlib-nav-toggle colorlib-nav-toggle"><i></i></a>
		<aside id="colorlib-aside" role="complementary" class="js-fullheight text-center">
			<h1 id="colorlib-logo"><a href="Index.asp"><img style="width: 50%;" src="images/Logo/logo_transparent.png"></a></h1>
			<nav id="colorlib-main-menu" role="navigation">
				<div class="desc">
               <h2 class="subheading">Workspace In Your Life</h2>
               <h1 class="mb-4">Zulu<span>.</span></h1>
               <p class="mb-4">Zulu helps you manage and operate everything in a fun way. People find an idea at a moment's notice and dream of putting it into practice, but most often forget it in a short time. Zulu helps you perform fine, perfect tasks next to you to prevent this unfortunate event.</p>
               <p><a href="#" class="btn-custom">More About Zulu <span class="ion-ios-arrow-forward"></span></a></p>
            </div>
			</nav>
			<div class="colorlib-footer">
				<p>Copyright &copy;<%= Left(Date(), 4) %> All rights reserved | <i class="icon-heart" aria-hidden="true"></i> by <a href="https://github.com/divashuthron" target="_blank">RH.BanYeoul</a>
				<ul>
					<li><a href="https://www.instagram.com/divashuthron/"><i class="icon-instagram"></i> Dev's instagram</a></li>
				</ul>
			</div>
		</aside>
		<div id="colorlib-main">
			<div class="hero-wrap js-fullheight" style="background-image: url('images/bg_2.jpg')" data-stellar-background-ratio="0.5">
				<div class="overlay"></div>
				<div class="js-fullheight d-flex justify-content-center align-items-center">
					<div class="col-md-8 text">
						<div class="row block-9">
                     <div class="col-md-6 mx-auto order-md-last pr-md-5">
                        <div class="desc text-center">
                           <h2 class="subheading text-dark">Workspace In Your Life</h2>
                           <h1 class="mb-4"><a href="Index.asp"><img style="width: 50%;" src="images/Logo/logo_transparent.png"></a></h1>
                        </div>
                        <form id="Login" class="justify-content-center align-items-center" action="/process/LoginProc.asp">
                           <div class="form-group">
                              <small><b><label class="text-dark" for="ID">아이디</label></b></small>
                              <input type="text" name="ID" class="form-control pl-3" placeholder="아이디">
                           </div>
                           <div class="form-group">
                              <small><b><label class="text-dark" for="Password">비밀번호</label></b></small>
                              <input type="password" name="Password" class="form-control pl-3" placeholder="비밀번호">
                           </div>
                           <div class="form-group text-center">
                              <Button type="button" id="Confirm" class="btn btn-primary py-3 px-5 mt-4">로그인</button>
                           </div>
                        </form>
                     </div>
                  </div>
					</div>
				</div>
			</div>
         <div id="ftco-loader" class="show fullscreen">
            <svg class="circular" width="48px" height="48px">
               <circle class="path-bg" cx="24" cy="24" r="22" fill="none" stroke-width="4" stroke="#eeeeee"/>
               <circle class="path" cx="24" cy="24" r="22" fill="none" stroke-width="4" stroke-miterlimit="10" stroke="#F96D00"/>
            </svg>
         </div>
<!-- #Include Virtual = "/Include/Bottom.asp" -->
<script>
      $('#Confirm').click(function() {
         var objOpt = {'url':'','param':'','dataType':'xml','before':'','success':'$.ConfirmResult(datas)','complete':'','clear':'','reset':''};
         objOpt["url"] = "/process/LoginProc.asp";
         $.Ajax4Form("#Login", objOpt);
         $("#Login").submit();
      });

      $.ConfirmResult = function(datas) {
         alert("?");
      }
</script>