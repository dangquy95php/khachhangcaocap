<!DOCTYPE html>
<html lang="en">

<head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title>@yield('title')</title>
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <meta name="csrf-token" content="{{ csrf_token() }}">
    <meta name="zalo-platform-site-verification" content="MjM41Qtg21DS_wWgl_LwSZ_qmNNJX18hDpa" />
    <link href="{{ asset('assets/img/favicon.png') }}" rel="icon">
    <link href="{{ asset('assets/img/apple-touch-icon.png') }}" rel="apple-touch-icon">
    <!-- Google Fonts -->
    <link href="https://fonts.gstatic.com" rel="preconnect">
    <link
        href="https://fonts.googleapis.com/css?family=Open+Sans:300,300i,400,400i,600,600i,700,700i|Nunito:300,300i,400,400i,600,600i,700,700i|Poppins:300,300i,400,400i,500,500i,600,600i,700,700i"
        rel="stylesheet">

    <!-- Vendor CSS Files -->
    <link href="{{ asset('assets/vendor/bootstrap/css/bootstrap.min.css') }}" rel="stylesheet">
    <link href="{{ asset('assets/vendor/bootstrap-icons/bootstrap-icons.css') }}" rel="stylesheet">
    <link href="{{ asset('assets/vendor/boxicons/css/boxicons.min.css') }}" rel="stylesheet">
    <link href="{{ asset('assets/vendor/quill/quill.snow.css') }}" rel="stylesheet">
    <link href="{{ asset('assets/vendor/quill/quill.bubble.css') }}" rel="stylesheet">
    <link href="{{ asset('assets/vendor/remixicon/remixicon.css') }}" rel="stylesheet">
    <link href="{{ asset('assets/vendor/simple-datatables/style.css') }}" rel="stylesheet">
    <link href="{{ asset('css/toastr.min.css') }}" rel="stylesheet" type="text/css">

    <!-- Template Main CSS File -->
    <link href="{{ asset('assets/css/style.css') }}" rel="stylesheet">
   
</head>

<body>
    <main>
        <div class="card">
            <h5 class="card-header text-dark">
                KHẢO SÁT KHÁCH HÀNG SỬ DỤNG PHẦN MỀM QUẢN LÝ NHÂN VIÊN SALES
            </h5>
            <form action="" method="post">
                @csrf 
                <div class="card-body">
                    <div class="card">
                        <div class="card-header text-primary">
                          Câu hỏi 1: Phần mềm quản lý nhân viên sales giao diện trải nghiệm có tốt hơn phần mềm cũ của bạn không?
                        </div>
                        <div class="card-body">
                            <div class="form-check pt-3">
                                <input class="form-check-input" type="radio" @if(old('question1')=="Tốt hơn" ) checked @endif name="question1" id="exampleRadios1" value="Tốt hơn">
                                <label class="form-check-label" for="exampleRadios1">Tốt hơn</label>
                            </div>
                            <div class="form-check">
                                <input class="form-check-input" type="radio" @if(old('question1')=="Không tốt hơn" ) checked @endif name="question1" id="exampleRadios2" value="Không tốt hơn">
                                <label class="form-check-label" for="exampleRadios2">Không tốt hơn</label>
                            </div>
                            <div class="form-check">
                                <input class="form-check-input" type="radio" @if(old('question1')=="Bình thường" ) checked @endif name="question1" id="exampleRadios3" value="Bình thường">
                                <label class="form-check-label" for="exampleRadios3">Bình thường</label>
                            </div>
                        </div>
                        @include('_partials.alert', ['field' => 'question1'])
    
                        <div class="card-header text-primary">
                            Câu hỏi 2: Tốc độ load dữ liệu có nhanh không?
                        </div>
                        <div class="card-body">
                            <div class="form-check pt-3">
                                <input class="form-check-input" type="radio" name="question2" @if(old('question2')=="Rất Nhanh" ) checked @endif id="exampleRadios2_1" value="Rất Nhanh">
                                <label class="form-check-label" for="exampleRadios2_1">Rất Nhanh</label>
                            </div>
                            <div class="form-check">
                                <input class="form-check-input" type="radio" name="question2" @if(old('question2')=="Nhanh" ) checked @endif id="exampleRadios2_3" value="Nhanh">
                                <label class="form-check-label" for="exampleRadios2_3">Nhanh</label>
                            </div>
                            <div class="form-check">
                                <input class="form-check-input" type="radio" name="question2"  @if(old('question2')=="Bình thường" ) checked @endif id="exampleRadios2_2" value="Bình thường">
                                <label class="form-check-label" for="exampleRadios2_2">Bình thường</label>
                            </div>
                            <div class="form-check">
                                <input class="form-check-input" type="radio" name="question2" @if(old('question2')=="Chậm" ) checked @endif id="exampleRadios2_3" value="Chậm">
                                <label class="form-check-label" for="exampleRadios2_3">Chậm</label>
                            </div>
                            <div class="form-check">
                                <input class="form-check-input" type="radio" name="question2" @if(old('question2')=="Rất Chậm" ) checked @endif id="exampleRadios2_4" value="Rất chậm">
                                <label class="form-check-label" for="exampleRadios2_4">Rất chậm</label>
                            </div>
                        </div>
                        @include('_partials.alert', ['field' => 'question2'])
    
                        <div class="card-header text-primary">
                            Câu hỏi 3: Có khó khăn trong việc sử dụng phần mềm không?
                        </div>
                        <div class="card-body">
                            <div class="form-check pt-3">
                                <input class="form-check-input" type="radio" name="question3" @if(old('question3')=="Không" ) checked @endif id="exampleRadios3_1" value="Không">
                                <label class="form-check-label" for="exampleRadios3_1">Không</label>
                            </div>
                            <div class="form-check">
                                <input class="form-check-input" type="radio" name="question3" @if(old('question3')=="Có" ) checked @endif id="exampleRadios3_2" value="Có">
                                <label class="form-check-label" for="exampleRadios3_2">Có</label>
                            </div>
                        </div>
                        @include('_partials.alert', ['field' => 'question3'])

                        <div class="card-header text-primary">
                            Câu hỏi 4: Bạn đánh giá trải nghiệm với phần mềm của chúng tôi như thế nào?
                        </div>
                        <div class="card-body">
                            <div class="form-check pt-3">
                                <input class="form-check-input" @if(old('question4')=="Rất hài lòng" ) checked @endif type="radio" name="question4" id="exampleRadios4_1" value="Rất hài lòng">
                                <label class="form-check-label" for="exampleRadios4_1">Rất hài lòng</label>
                            </div>
                            <div class="form-check">
                                <input class="form-check-input" @if(old('question4')=="Hài lòng" ) checked @endif type="radio" name="question4" id="exampleRadios4_2" value="Hài lòng">
                                <label class="form-check-label" for="exampleRadios4_2">Hài lòng </label>
                            </div>
                            <div class="form-check">
                                <input class="form-check-input" @if(old('question4')=="Chưa hài lòng" ) checked @endif type="radio" name="question4" id="exampleRadio4_3" value="Chưa hài lòng">
                                <label class="form-check-label" for="exampleRadio4_3">Chưa hài lòng </label>
                            </div>
                        </div>
                        @include('_partials.alert', ['field' => 'question4'])

                        <div class="card-header text-primary">
                            Câu hỏi 5: Bạn có muốn nâng cấp thêm nhiều chức năng cho phần mềm này trong tương lai không?
                        </div>
                        <div class="card-body">
                            <div class="form-check pt-3">
                                <input class="form-check-input" type="radio" name="question5" id="exampleRadios5_1" value="Muốn">
                                <label class="form-check-label" for="exampleRadios5_1">Muốn</label>
                            </div>
                            <div class="form-check">
                                <input class="form-check-input" type="radio" name="question5" id="exampleRadios5_2" value="Không">
                                <label class="form-check-label" for="exampleRadios5_2">Không </label>
                            </div>
                            <div class="form-check">
                                <input class="form-check-input" type="radio" name="question5" id="exampleRadio5_3" value="Có dự định nhưng suy nghĩ sau">
                                <label class="form-check-label" for="exampleRadio5_3">Có dự định nhưng suy nghĩ sau</label>
                            </div>
                        </div>
                        @include('_partials.alert', ['field' => 'question5'])

                        <div class="card-header text-primary">
                            Câu hỏi 6: Bạn yêu thích điều gì nhất về phần mềm nhân viên sales chúng tôi?
                        </div>
                        <div class="card-body">
                            <div class="form-group">
                                <textarea class="form-control" name="question6" id="exampleFormControlTextarea1" rows="3">{{ old('question6') }}</textarea>
                            </div>
                        </div>
                        @include('_partials.alert', ['field' => 'question6'])

                        <div class="card-header text-primary">
                            Câu hỏi 7: Bạn có cảm thấy thoải mái khi mua phần mềm của chúng tôi?
                        </div>
                        <div class="card-body">
                            <div class="form-group">
                                <textarea class="form-control" name="question7" id="exampleFormControlTextarea1" rows="3">{{ old('question7') }}</textarea>
                            </div>
                        </div>
                        @include('_partials.alert', ['field' => 'question7'])

                        <div class="card-header text-primary">
                            Câu hỏi 8: Làm thế nào để chúng tôi có thể phục vụ bạn tốt hơn?
                        </div>
                        <div class="card-body">
                            <div class="form-group">
                                <textarea class="form-control" name="question8" id="exampleFormControlTextarea1" rows="3">{{ old('question8') }}</textarea>
                            </div>
                        </div>
                        @include('_partials.alert', ['field' => 'question8'])
                    </div>
                    <button type="submit" class="btn btn-success  btn-lg">Gửi</button>
                </div>
            </form>
          </div>
    </main>

    <footer class="footer">
        <div class="copyright">
            © Copyright <strong><span>PxwebShop</span></strong>. Đã đăng ký bản quyền
        </div>
        <div class="credits">
            <!-- All the links in the footer should remain intact. -->
            <!-- You can delete the links only if you purchased the pro version. -->
            <!-- Licensing information: https://bootstrapmade.com/license/ -->
            <!-- Purchase the pro version with working PHP/AJAX contact form: https://bootstrapmade.com/nice-admin-bootstrap-admin-html-template/ -->
            Thiết kế bởi <a href="https://pxwebshop.com/">Công Ty TNHH PxwebShop</a>
        </div>
    </footer>

    <!-- Vendor JS Files -->
    <script src="{{ asset('assets/vendor/apexcharts/apexcharts.min.js') }}"></script>
    <script src="{{ asset('assets/vendor/bootstrap/js/bootstrap.bundle.min.js') }}"></script>
    <script src="{{ asset('assets/vendor/chart.js/chart.min.js') }}"></script>
    <script src="{{ asset('assets/vendor/echarts/echarts.min.js') }}"></script>
    <script src="{{ asset('assets/vendor/quill/quill.min.js') }}"></script>
    <script src="{{ asset('assets/vendor/simple-datatables/simple-datatables.js') }}"></script>
    <script src="{{ asset('assets/vendor/tinymce/tinymce.min.js') }}"></script>
    <script src="{{ asset('assets/vendor/php-email-form/validate.js') }}"></script>
    <!-- Template Main JS File -->
    <script src="{{ asset('assets/js/main.js') }}"></script>

    <!-- Jquery Slim JS -->
    <script src="{{ asset('js/jquery.min.js') }} "></script>

    <script src="{{ asset('js/jquery-ui.min.js') }} "></script>

    <svg id="SvgjsSvg1145" width="2" height="0" xmlns="http://www.w3.org/2000/svg" version="1.1"
        xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:svgjs="http://svgjs.dev"
        style="overflow: hidden; top: -100%; left: -100%; position: absolute; opacity: 0;">
        <defs id="SvgjsDefs1146"></defs>
        <polyline id="SvgjsPolyline1147" points="0,0"></polyline>
        <path id="SvgjsPath1148"
            d="M-1 270.2L-1 270.2C-1 270.2 176.9170673076923 270.2 176.9170673076923 270.2C176.9170673076923 270.2 294.86177884615387 270.2 294.86177884615387 270.2C294.86177884615387 270.2 412.80649038461536 270.2 412.80649038461536 270.2C412.80649038461536 270.2 530.7512019230769 270.2 530.7512019230769 270.2C530.7512019230769 270.2 648.6959134615385 270.2 648.6959134615385 270.2C648.6959134615385 270.2 766.640625 270.2 766.640625 270.2C766.640625 270.2 766.640625 270.2 766.640625 270.2 ">
        </path>
    </svg>

    <script src="{{ asset('js/toastr.min.js') }} "></script>


    {!! Toastr::message() !!}

</body>

</html>
