<!DOCTYPE html>
<html lang="en">

<head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title>Đăng Nhập</title>
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <link href="{{ asset('assets/img/favicon.png')}}" rel="icon">
    <link href="{{ asset('assets/img/apple-touch-icon.png')}}" rel="apple-touch-icon">

    <!-- Google Fonts -->
    <link href="https://fonts.gstatic.com" rel="preconnect">
    <link href="https://fonts.googleapis.com/css?family=Open+Sans:300,300i,400,400i,600,600i,700,700i|Nunito:300,300i,400,400i,600,600i,700,700i|Poppins:300,300i,400,400i,500,500i,600,600i,700,700i" rel="stylesheet">

    <!-- Vendor CSS Files -->
    <link href="{{ asset('assets/vendor/bootstrap/css/bootstrap.min.css')}}" rel="stylesheet">
    <link href="{{ asset('assets/vendor/bootstrap-icons/bootstrap-icons.css')}}" rel="stylesheet">
    <link href="{{ asset('assets/vendor/boxicons/css/boxicons.min.css')}}" rel="stylesheet">
    <link href="{{ asset('css/toastr.min.css') }}" rel="stylesheet" type="text/css" >
      <!-- Template Main CSS File -->
      <link href="{{ asset('assets/css/style.css')}}" rel="stylesheet">
</head>

<body>

<main>
   <div class="container">
      <section class="section register min-vh-100 d-flex flex-column align-items-center justify-content-center py-4">
         <div class="container">
            <div class="row justify-content-center">
               <div class="col-lg-4 col-md-6 d-flex flex-column align-items-center justify-content-center">
                  <div class="d-flex justify-content-center py-4">
                     <a href="index.html" class="logo d-flex align-items-center w-auto">
                     <img src="{{asset('assets/img/logo.png')}}" alt="">
                     <span class="d-none d-lg-block">NHÂN VIÊN SALES</span>
                     </a>
                  </div>
                  <!-- End Logo -->
                  <div class="card mb-3">
                     <div class="card-body pb-4">
                        <div class="pt-4 pb-2">
                           <h5 class="card-title text-center pb-0 fs-4">Đăng Nhập</h5>
                           <p class="text-center small">Nhập tên người dùng và mật khẩu để đăng nhập</p>
                        </div>
                        <form action="{{route('post_login')}}" class="row g-3 needs-validation" method="POST" novalidate>
                           @csrf
                           <div class="col-12">
                              <label for="yourUsername" class="form-label">Tên người dùng</label>
                              <input type="text" value="{{ old('username') }}"  name="username" class="form-control" id="yourUsername" required>
                              <div class="invalid-feedback">Please enter your username.</div>
                              @include('_partials.alert', ['field' => 'username'])
                           </div>
                           <div class="col-12">
                              <label for="yourPassword" class="form-label">Mật khẩu</label>
                              <input type="password" value="{{ old('password') }}" name="password" class="form-control" id="yourPassword" required>
                              <div class="invalid-feedback">Please enter your password!</div>
                              @include('_partials.alert', ['field' => 'password'])
                           </div>
                           <div class="col-12">
                              <div class="form-check">
                                 <input class="form-check-input" type="checkbox" name="remember" value="true" id="rememberMe">
                                 <label class="form-check-label" for="rememberMe">Nhớ mật khẩu</label>
                              </div>
                           </div>
                           <div class="col-12">
                              <button class="btn btn-primary w-100" type="submit">Đăng Nhập</button>
                           </div>
                           <div class="col-12">
                              <p class="small mb-0">Người tạo phần mềm: <b>Đặng Quý (0964.944.719)</b></p>
                            </div>
                           <!-- <div class="col-12">
                              <p class="small mb-0">Tôi chưa có tài khoản? <a href="{{route('create_account')}}">Tạo tài khoản mới!</a></p>
                           </div> -->
                        </form>
                     </div>
                  </div>
                  <div class="credits">
                     <!-- All the links in the footer should remain intact. -->
                     <!-- You can delete the links only if you purchased the pro version. -->
                     <!-- Licensing information: https://bootstrapmade.com/license/ -->
                     <!-- Purchase the pro version with working PHP/AJAX contact form: https://bootstrapmade.com/nice-admin-bootstrap-admin-html-template/ -->
                    Thiết kế bởi <a href="https://pxwebshop.com/">Công Ty TNHH PxwebShop</a>
                  </div>
               </div>
            </div>
         </div>
      </section>
   </div>
</main>
<!-- End #main -->

<!-- Jquery Slim JS -->
<script src="{{ asset('js/jquery.min.js')}} "></script>

<script src="{{ asset('js/jquery-ui.min.js')}} "></script>

<!-- Vendor JS Files -->
<script src="{{asset('assets/vendor/bootstrap/js/bootstrap.bundle.min.js')}}"></script>
<script src="{{asset('assets/vendor/php-email-form/validate.js')}}"></script>

<script src="{{ asset('js/toastr.min.js')}} "></script>

{!! Toastr::message() !!}

</body>

</html>
