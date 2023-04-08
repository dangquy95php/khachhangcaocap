<!DOCTYPE html>
<html lang="en">

<head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title>Đăng Ký Tài Khoản Nhân Viên</title>
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
                     <span class="d-none d-lg-block">NHÂN VIÊN Sales</span>
                     </a>
                  </div>
                  <!-- End Logo -->
                  <div class="card mb-3">
                     <div class="card-body">
                        <div class="pt-4 pb-2">
                           <h5 class="card-title text-center pb-0 fs-4">Tạo Tài Khoản</h5>
                           <p class="text-center small">Nhập thông tin chi tiết của bạn để tạo tài khoản</p>
                        </div>
                        <form class="row g-3 needs-validation" method="post" novalidate>
                            @csrf
                           <div class="col-12">
                              <label for="yourName" class="form-label">Tên của bạn</label>
                              <input type="text" value="{{old('name', @$data->name)}}" name="name" class="form-control" id="yourName">
                           </div>
                           <div class="col-12">
                              <label for="yourEmail" class="form-label">Email</label>
                              <input type="email" value="{{old('email', @$data->email)}}" name="email" class="form-control" id="yourEmail">
                           </div>
                           <div class="col-12">
                              <label for="yourUsername" class="form-label">Tên đăng nhập <span class="badge rounded-pill bg-warning text-dark">Viết không dấu</span></label>
                                 <input type="text" {{ \Request::route()->getName() == 'edit_account' ? 'disabled' : '' }} value="{{old('username', @$data->username)}}" name="username" class="form-control" id="yourUsername" required>
                                 @include('_partials.alert', ['field' => 'username'])
                           </div>
                           <div class="col-12">
                              <label for="yourPassword" class="form-label">Mật khẩu</label>
                              <input type="password" name="password" class="form-control" id="yourPassword" required>
                              @include('_partials.alert', ['field' => 'password'])
                           </div>

                           <div class="col-12">
                              <label for="yourPassword1" class="form-label">Nhập lại mật khẩu</label>
                              <input type="password" name="password_confirmation" class="form-control" id="yourPassword1" required>
                              @include('_partials.alert', ['field' => 'password_confirmation'])
                           </div>
                           @if(\Request::route()->getName() == 'edit_account')
                           <div class="col-12">
                                <div class="form-check">
                                    <input class="form-check-input" name="check_password" type="checkbox" id="flexCheckChecked">
                                    <label class="form-check-label" for="flexCheckChecked">
                                       Tôi không muốn thay đổi mật khẩu cũ
                                    </label>
                                </div>
                           </div>
                           @endif

                           <div class="col-12">
                                <label for="yourPassword" class="form-label">Cấp quyền truy cập</label>
                                <select class="form-select" name="role" aria-label="Default select example">
                                    <option vlue="">Chọn quyền truy cập</option>
                                    @foreach(App\Models\User::getRole() as $role)
                                        <option
                                            {{ old('role') == $role ? "selected" : "" }}

                                            @if($role == 1 && @$data->role == 1)
                                                selected
                                            @endif
                                            @if($role == 2 && @$data->role == 2)
                                                selected
                                            @endif

                                            value="{{$role}}">
                                            @if($role == 1)
                                                Người dùng
                                            @endif
                                            @if($role == 2)
                                                Admin
                                            @endif
                                        </option>
                                    @endforeach
                                </select>
                                @include('_partials.alert', ['field' => 'status'])
                           </div>

                            <div class="col-12">
                                <label for="yourPassword" class="form-label">Chọn trạng thái tài khoản</label>
                                <select class="form-select" name="status" aria-label="Default select example">
                                    <option vlue="">Vui lòng chọn</option>
                                    @foreach(App\Models\User::getStatus() as $status)
                                        <option
                                            @if($status == 0 && @$data->status == 0)
                                                selected
                                            @endif
                                            @if($status == 1 && @$data->status == 1)
                                                selected
                                            @endif
                                            @if($status == 2 && @$data->status == 2)
                                                selected
                                            @endif

                                            {{ old('status') == $status ? "selected" : "" }}

                                            value="{{$status}}">
                                            @if($status == 1)
                                                Đang hoạt động
                                            @endif
                                            @if($status == 2)
                                                Không hoạt động
                                            @endif
                                            @if($status == 0)
                                                Chưa kích hoạt
                                            @endif
                                        </option>
                                    @endforeach
                                </select>
                                @include('_partials.alert', ['field' => 'status'])
                           </div>
                           <div class="col-12 mt-4">
                              <button class="btn btn-primary w-100" type="submit">Đăng Ký</button>
                           </div>
                           @if(\Request::route()->getName() !== 'create_account')
                           <div class="col-12">
                              <p class="small mb-0">Tôi đã có tài khoản? <a href="{{route('login')}}">Đăng Nhập</a></p>
                           </div>
                           @endif
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
<a href="#" class="back-to-top d-flex align-items-center justify-content-center"><i class="bi bi-arrow-up-short"></i></a>

<!-- Vendor JS Files -->
<script src="{{asset('assets/vendor/bootstrap/js/bootstrap.bundle.min.js')}}"></script>
<script src="{{asset('assets/vendor/php-email-form/validate.js')}}"></script>
