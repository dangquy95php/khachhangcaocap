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
    <style>
        .ql-editor {
            min-height: 100px;
        }
        .is-copy {
            right: 5px;
            top: 7px;
            cursor: pointer;
        }
        .is-clipboad {
            opacity: 0;
        }
    </style>
</head>

<body>
    <header id="header" class="header fixed-top d-flex align-items-center">
        <div class="d-flex align-items-center justify-content-between">
            <a href="/" class="logo d-flex align-items-center">
                <img src="{{ asset('assets/img/logo.png') }}" alt="">
                <span class="d-none d-lg-block">Nhân Viên Sales</span>
            </a>
            {{\Request()->route()->getPrefix() == 'admin' ? '<i class="bi bi-list toggle-sidebar-btn"></i>' : '' }}
        </div>
        <!-- End Logo -->

        <!-- End Search Bar -->
        <nav class="header-nav ms-auto">
            <ul class="d-flex align-items-center">
                <li class="nav-item d-block d-none">
                    <a class="nav-link nav-icon search-bar-toggle " href="#">
                        <i class="bi bi-search"></i>
                    </a>
                </li>
                <!-- End Search Icon-->

                <!-- End Messages Nav -->
                <li class="nav-item dropdown pe-3">
                    <a class="nav-link nav-profile d-flex align-items-center pe-0" href="#" data-bs-toggle="dropdown">
                        <span class="d-none d-md-block dropdown-toggle ps-2">{{ Auth::user()->username }}</span>
                    </a><!-- End Profile Iamge Icon -->
                    <ul class="dropdown-menu dropdown-menu-end dropdown-menu-arrow profile">
                        @if (Auth::user()->role == 2)
                            <li>
                                <a class="dropdown-item d-flex align-items-center" href="{{ route('dashboard') }}">
                                    <i class="ri-admin-line"></i>
                                    <span>Trang Quản Trị</span>
                                </a>
                            </li>
                            <li>
                                <hr class="dropdown-divider">
                            </li>
                        @endif
                        <li>
                            <a class="dropdown-item d-flex align-items-center" href="{{ route('logout') }}">
                                <i class="bi bi-box-arrow-right"></i>
                                <span>Đăng Xuất</span>
                            </a>
                        </li>
                    </ul>
                    <!-- End Profile Dropdown Items -->
                </li>
                <!-- End Profile Nav -->
            </ul>
        </nav>
        <!-- End Icons Navigation -->
    </header>

    <!-- End Sidebar-->

    <main style="margin-top:80px;">

        <div class="card">
            <div class="card-body px-3">
                <h5 class="card-title mb-0">PHẦN MỀM CHĂM SÓC KHÁCH HÀNG</h5>

                <div class="card pt-3 mb-3">
                    <div class="card-body p-md-2">
                        <!-- Vertical Form -->
                        <form method="POST" action="" class="row g-3" id="customerForm">
                            <input type="text" class="d-none" id="customer_id" name="id" value="{{ @$customer->id }}" />
                            @csrf
                            <div class="col-lg-12">
                                <div class="row">
                                    <div class="col-lg-5 col-md-8 col-sm-12 pe-md-1">
                                        <div class="row g-3">
                                            <div class="col-sm-4 pe-md-2">
                                                <label for="inputPassword4" class="form-label text-danger"><b>Số Hợp Đồng</b></label>
                                                <div class="position-relative is-copy-wrap">
                                                    <input type="text" disabled name="id_contract" value="{{ @$customer->so_hop_dong }}" id="id_contract" class="form-control" id="inputPassword4">
                                                    <span class="position-absolute is-copy"><span class="is-clipboad">{{ @$customer->so_hop_dong }}</span><i class="bi bi-clipboard is-status-clip"></i></span>
                                                </div>
                                            </div>
                                            <div class="col-sm-4">
                                                <label for="inputEmail4" class="form-label"><b>Ngày Tham Gia</b></label>
                                                <input type="text" class="form-control" value="{{ @$customer->ngay_tham_gia }}" id="ngay_tham_gia" name="ngay_tham_gia">
                                            </div>
                                            <div class="col-sm-4">
                                                <label for="inputNanme4" class="form-label"><b>Năm Đáo Hạn</b></label>
                                                <input id="nam_dao_han" name="nam_dao_han" class="form-control" type="text" value="{{ @$customer->nam_dao_han }}" />
                                            </div>
                                            <div class="col-sm-4">
                                                <label for="inputEmail4" class="form-label"><b>Số Tiền</b></label>
                                                <input type="text" class="form-control" value="{{ is_numeric(@$customer->menh_gia) ? number_format(@$customer->menh_gia + 50000000) : @$customer->menh_gia }}" id="menh_gia" name="money">
                                            </div>

                                            <div class="col-sm-4">
                                                <label class="form-label"><b>VP/Bank</b></label>
                                                <input type="text" class="form-control" value="{{ @$customer->vpbank }}" id="vpbank" name="vpbank">
                                            </div>
                                            <div class="col-sm-4">
                                                <label class="form-label"><b>CV</b></label>
                                                <input type="text" class="form-control" value="{{ @$customer->cv }}" id="cv" name="cv">
                                            </div>
                                            <div class="col-md-12 col-sm-12">
                                                <label for="inputNanme4" class="form-label"><b>Họ Và Tên</b></label>
                                                <input type="text" value="{{ @$customer->ten_kh }}" name="last_name" class="form-control" id="ten_kh">
                                            </div>
                                            <div class="col-md-7 col-sm-12">
                                                <label for="inputNanme4" class="form-label"><b>CCCD</b></label>
                                                <input type="text" name="cccd" value="{{ @$customer->cccd }}" class="form-control" id="cccd">
                                            </div>
                                            <div class="col-md-5 col-sm-12">
                                                <label for="inputNanme4" class="form-label"><b>Số Điện Thoại</b></label>
                                                <input type="text" name="phone" value="{{ @$customer->dien_thoai }}" class="form-control" id="dien_thoai">
                                            </div>
                                            <div class="col-12">
                                            <label for="inputPassword4" class="form-label"><b>Địa Chỉ</b></label>
                                            <input type="text" value="{{ @$customer->dia_chi_cu_the }}"
                                                name="address_full" id="dia_chi_cu_the" class="form-control"
                                                id="inputPassword4">
                                            </div>
                                            <div class="col-md-2 col-sm-6 pe-0">
                                            <label for="inputPassword4" class="form-label"><b>Tuổi</b></label>
                                            <input type="number" value="{{ @$customer->tuoi }}" name="age"
                                                id="tuoi" min="1" max="200"
                                                onkeyup="if(parseInt(this.value) > 200 || parseInt(this.value) < 1){ this.value = ''; return false; }"
                                                class="form-control" id="inputPassword4">
                                            </div>
                                            <div class="col-md-3 col-sm-6">
                                            <label for="inputEmail4" class="form-label"><b>Giới Tính</b></label>
                                            <select name="sex" id="gioi_tinh" class="form-select">
                                                <option selected="">Chọn giới tính...</option>
                                                <option {{ @$customer->gioi_tinh === 'M' ? 'selected' : '' }}
                                                value="M">Nam</option>
                                                <option {{ @$customer->gioi_tinh === 'F' ? 'selected' : '' }}
                                                value="F">Nữ</option>
                                            </select>
                                            </div>
                                            <div class="col-md-7 col-sm-12">
                                            <label for="inputEmail4" class="form-label"><b>Nguồn Dữ
                                            Liệu</b></label>
                                            <select name="area_name" id="data_area_id" class="form-select">
                                                <option value="">Chọn nguồn dữ liệu</option>
                                                @foreach ($areas->areas as $area)
                                                    @if (!empty(old('area_name')))
                                                        <option {{ old('area_name') == @$area->id ? 'selected' : '' }} value="{{ @$area->id }}">{{ @$area->name }}</option>
                                                    @else
                                                        <option {{ @$area->id == @$customer->area_id ? 'selected' : ''}} value="{{ @$area->id }}">{{ @$area->name }}</option>
                                                    @endif
                                                @endforeach
                                            </select>
                                            </div>
                                            @error('area_name')
                                            <div class="text-danger text-end mt-1"> {{ $message }}</div>
                                            @enderror
                                        </div>
                                    </div>
                                    <div class="col-lg-4 col-md-4 col-sm-6 px-md-1 mt-md-0 mt-sm-3">
                                        <label for="inputEmail4" class="form-label"><b>Kết Quả Gọi</b></label>
                                        <ul class="list-group is-result">
                                            <li class="list-group-item">
                                            <div class="form-check">
                                                <input class="form-check-input" type="radio" name="type_call" value="0" id="gridRadios0"
                                                {{ old('type_call') !== null && old('type_call', @$customer->type_call) == 0 ? 'checked' : '' }}>
                                                <label class="form-check-label" for="gridRadios0">
                                                Đã hẹn
                                                </label>
                                            </div>
                                            </li>
                                            <li class="list-group-item">
                                            <div class="form-check">
                                                <input class="form-check-input" type="radio" name="type_call"
                                                value="1" id="gridRadios1"
                                                {{ old('type_call', @$customer->type_call) == 1 ? 'checked' : '' }}>
                                                <label class="form-check-label" for="gridRadios1">
                                                Đại lý vẫn chăm sóc
                                                </label>
                                            </div>
                                            </li>
                                            <li class="list-group-item">
                                            <div class="form-check">
                                                <input class="form-check-input" type="radio" name="type_call"
                                                value="2" id="gridRadios2"
                                                {{ old('type_call', @$customer->type_call) == 2 ? 'checked' : '' }}>
                                                <label class="form-check-label" for="gridRadios2">
                                                Khách hàng ít tiền
                                                </label>
                                            </div>
                                            </li>
                                            <li class="list-group-item">
                                            <div class="form-check">
                                                <input class="form-check-input" type="radio" name="type_call"
                                                value="3" id="gridRadios3"
                                                {{ old('type_call', @$customer->type_call) == 3 ? 'checked' : '' }}>
                                                <label class="form-check-label" for="gridRadios3">
                                                Khách hàng suy nghĩ, gọi lại sau
                                                </label>
                                            </div>
                                            </li>
                                            <li class="list-group-item">
                                            <div class="form-check">
                                                <input class="form-check-input" type="radio" name="type_call" value="4" id="gridRadios4"
                                                {{ old('type_call', @$customer->type_call) == 4 ? 'checked' : '' }}>
                                                <label class="form-check-label" for="gridRadios4">
                                                KNM / Bận / Tắt máy / Sai số / Đổi số
                                                </label>
                                            </div>
                                            </li>
                                            <li class="list-group-item">
                                            <div class="form-check">
                                                <input class="form-check-input" type="radio" name="type_call"
                                                value="5" id="gridRadios5"
                                                {{ old('type_call', @$customer->type_call) == 5 ? 'checked' : '' }}>
                                                <label class="form-check-label" for="gridRadios5">
                                                Hợp đồng Hủy / Đáo hạn / Hoàn trả giá trị
                                                </label>
                                            </div>
                                            </li>
                                            <li class="list-group-item">
                                            <div class="form-check">
                                                <input class="form-check-input" type="radio" name="type_call"
                                                value="6" id="gridRadios6"
                                                {{ old('type_call', @$customer->type_call) == 6 ? 'checked' : '' }}>
                                                <label class="form-check-label" for="gridRadios6">
                                                Đã tham gia hợp đồng mới / Không tham gia
                                                </label>
                                            </div>
                                            </li>
                                            <li class="list-group-item">
                                            <div class="form-check">
                                                <input class="form-check-input" type="radio" name="type_call"
                                                value="7" id="gridRadios7"
                                                {{ old('type_call', @$customer->type_call) == 7 ? 'checked' : '' }}>
                                                <label class="form-check-label" for="gridRadios7">
                                                Khác
                                                </label>
                                            </div>
                                            </li>
                                        </ul>
                                        @include('_partials.alert', ['field' => 'type_call'])
                                    </div>

                                    <div class="col-lg-3 col-md-12 col-sm-6 mt-lg-0 mt-md-4 mt-sm-3 px-md-1">
                                        <div class="col-12 col-sm-12">
                                            <!-- <textarea class="form-control" placeholder="Ghi chú cụ thể thông tin khách hàng" id="floatingTextarea" rows="5"></textarea> -->
                                            <div class="card mb-2">
                                                <div class="card-body">
                                                    <h5 class="card-title pt-3 pb-0"><b>Ghi Chú</b></h5>
                                                    <textarea name="comment" autofocus id="comment" value="{{ @$customer->comment }}" class="form-control" rows="10">{{ old('comment', @$customer->comment) }}</textarea>
                                                    <!-- End Quill Editor default -->
                                                </div>
                                            </div>
                                        </div>
                                        <div class="col-lg-12 col-sm-6">
                                            <!-- <button type="button" class="btn btn-danger"><i class="bi bi-exclamation-octagon"></i><span class="ps-2">Reset Dữ liệu</span></button> -->
                                            <button type="submit" class="d-none btn-save btn btn-success" name="action" value="save">
                                                <i class="bi bi-check-circle"></i><span class="ps-2">Cập Nhật Dữ Liệu</span>
                                            </button>

                                            <button type="submit" class="btn-next btn btn-outline-primary" name="action"
                                                value="next">
                                                <span class="pe-1">Gọi Tiếp Theo</span>
                                                <i class="bi bi-chevron-double-right"></i>
                                            </button>
                                        </div>
                                    </div>
                                </div>
                             </div>
                        </form><!-- Vertical Form -->
                    </div>
                </div>
                <!-- Default Tabs -->
                <ul class="nav nav-tabs" id="myTab" role="tablist">
                    <li class="nav-item" role="presentation">
                        <button class="nav-link active" id="profile-tab" data-bs-toggle="tab" data-bs-target="#profile"
                            type="button" role="tab" aria-controls="profile" aria-selected="false">Danh sách vừa gọi</button>
                    </li>
                    <li class="nav-item" role="presentation">
                        <button class="nav-link" id="history-tab" data-bs-toggle="tab" data-bs-target="#history" type="button" role="tab" aria-controls="history" aria-selected="false">Lịch sử đã gọi</button>
                    </li>
                </ul>
                <div class="tab-content pt-2" id="myTabContent">
                      <!-- //data 2 -->
                    <div class="tab-pane show active" id="profile" role="tabpanel" aria-labelledby="profile-tab">
                        <div class="card mb-2 mt-md-2 mt-sm-3">
                            <div class="card-body pt-3 table-responsive">
                                {{ count($today) == 0 ? `<h5 class="text-center"><b>Dữ Liệu Chưa Có</b></h5>` : '' }}
                                <!-- Table with hoverable rows -->
                                <table class="table table-hover" style="min-width: 1000px;">
                                    <thead>
                                        <tr>
                                            <th scope="col">#</th>
                                            <th class="d-none" scope="col">Ngày Tham Gia</th>
                                            <th class="d-none" scope="col">Năm Đáo Hạn</th>
                                            <th scope="col">Số Hợp Đồng</th>
                                            <th class="d-none" scope="col">VP/Bank</th>
                                            <th class="d-none" scope="col">CV</th>
                                            <th style="min-width: 150px" scope="col">Tên</th>
                                            <th scope="col">Tuổi</th>
                                            <th scope="col">Giới Tính</th>
                                            <th scope="col">Số Điện Thoại</th>
                                            <th style="min-width:200px;" scope="col">Ghi Chú</th>
                                            <th scope="col">Mệnh Giá</th>
                                            <th style="width:200px;" scope="col">Kết Quả Cuộc Gọi</th>
                                            <th style="min-width: 250px" scope="col">Địa Chỉ</th>
                                            <th style="width:150px;" scope="col">Ngày Gọi</th>
                                            <th class="d-none" scope="col">Nguồn Dữ Liệu</th>
                                        </tr>
                                    </thead>
                                    <tbody>
                                        @php
                                        $i = count($today);
                                        @endphp
                                        @foreach ($today as $key => $data)
                                        <tr role="button" class="is-item-customer">
                                            <th scope="row">
                                                @if ($key == 0)
                                                    {{ $i }}
                                                @else
                                                    {{ $i }}
                                                @endif
                                            </th>
                                            <th class="d-none id" scope="row">{{ @$data->id }}</th>
                                            <td class="d-none ngay_tham_gia">{{ @$data->ngay_tham_gia }}</td>
                                            <td class="d-none nam_dao_han">{{ @$data->nam_dao_han }}</td>
                                            <td class="so_hop_dong">{{ $data->so_hop_dong }}</td>
                                            <td class="d-none vpbank">{{ $data->vpbank }}</td>
                                            <td class="d-none cv">{{ $data->cv }}</td>
                                            <td class="ten_kh"><b>{{ $data->ten_kh }}</b></td>
                                            <td class="tuoi">
                                                <span class="{{ ($data->tuoi > 50) ? 'badge bg-dark' : '' }}">{{ $data->tuoi }}</span>
                                            </td>
                                            <td class="gioi_tinh">
                                                @if($data->gioi_tinh === 'M')
                                                Nam
                                                @elseif($data->gioi_tinh === 'F')
                                                Nữ
                                            @endif
                                            </td>
                                            <td class="dien_thoai">{{ $data->dien_thoai }}</td>
                                            <td class="d-none cccd">{{ $data->cccd }}</td>
                                            <td class="comment">{{ $data->comment }}</td>
                                            <td class="menh_gia">{{ is_numeric(@$data->menh_gia) ? number_format(@$data->menh_gia + 50000000) : @$data->menh_gia }}</td>
                                            <td class="type_call">
                                                @foreach (\App\Models\Customer::getInforOption() as $key => $value)
                                                @if ($key == $data->type_call)
                                                <span data-id="{{ $key }}" class="badge {{ $key == 0 ? 'bg-danger' : 'bg-primary' }} ">{{ $value }}</span>
                                                @endif
                                                @endforeach
                                            </td>
                                            <td class="dia_chi_cu_the">{{ $data->dia_chi_cu_the }}</td>
                                            <td class="updated_at">{{ $data->updated_at }}</td>
                                            <td class="d-none area_name">{{ $data->name }}</td>
                                        </tr>
                                        @php
                                        $i--;
                                        @endphp
                                        @endforeach
                                    </tbody>
                                </table>
                                <!-- End Table with hoverable rows -->
                            </div>
                        </div>
                    </div><!-- End Default Tabs -->
                    <div class="tab-pane fade" id="history" role="tabpanel" aria-labelledby="history-tab">

                        <div class="col-12">

                            <div class="card mb-2 mt-md-2 mt-sm-3">
                                <div class="card-body pt-3 table-responsive">
                                    {{ count($history->histories) == 0 ? `<h5 class="text-center"><b>Dữ Liệu Chưa Có</b></h5>` : '' }}
                                    <!-- Table with hoverable rows -->
                                    <table class="table table-hover">
                                        <thead>
                                            <tr>
                                                <th scope="col">#</th>
                                                <th class="d-none" scope="col">Ngày Tham Gia</th>
                                                <th class="d-none" scope="col">Năm Đáo Hạn</th>
                                                <th scope="col">Số Hợp Đồng</th>
                                                <th class="d-none" scope="col">VP/Bank</th>
                                                <th class="d-none" scope="col">CV</th>
                                                <th style="min-width: 150px" scope="col">Tên</th>
                                                <th scope="col">Tuổi</th>
                                                <th scope="col">Giới Tính</th>
                                                <th scope="col">Số Điện Thoại</th>
                                                <th style="min-width:200px;" scope="col">Ghi Chú</th>
                                                <th scope="col">Mệnh Giá</th>
                                                <th style="width:200px;" scope="col">Kết Quả Cuộc Gọi</th>
                                                <th style="min-width: 250px" scope="col">Địa Chỉ</th>
                                                <th style="width:100px;" scope="col">Ngày Gọi</th>
                                                <th class="d-none" scope="col">Nguồn Dữ Liệu</th>
                                            </tr>
                                        </thead>
                                        <tbody>
                                            @php
                                            $j = $history->histories->total();
                                            if ($history->histories->currentPage() >= 2) {
                                                $j = $history->histories->total() - (($history->histories->currentPage() - 1) * $history->histories->perPage());
                                            }
                                            @endphp

                                            @foreach ($history->histories as $key => $data)
                                            <tr role="button" class="is-item-customer">
                                                <th scope="row">
                                                    @if ($key == 0 && $history->histories->currentPage() < 2)
                                                        {{ $j }}
                                                    @else
                                                        {{ $j }}
                                                    @endif
                                                </th>
                                                <th class="d-none id" scope="row">{{ @$data->id }}</th>
                                                <td class="d-none ngay_tham_gia">{{ @$data->ngay_tham_gia }}</td>
                                                <td class="d-none nam_dao_han">{{ @$data->nam_dao_han }}</td>
                                                <td class="so_hop_dong">{{ $data->so_hop_dong }}</td>
                                                <td class="d-none vpbank">{{ $data->vpbank }}</td>
                                                <td class="d-none cv">{{ $data->cv }}</td>
                                                <td class="ten_kh"><b>{{ $data->ten_kh }}</b></td>
                                                <td class="tuoi">
                                                    <span class="{{ ($data->tuoi > 50) ? 'badge bg-dark' : '' }}">{{ $data->tuoi }}</span>
                                                </td>
                                                <td class="gioi_tinh">
                                                @if($data->gioi_tinh === 'M')
                                                    Nam
                                                    @elseif($data->gioi_tinh === 'F')
                                                    Nữ
                                                @endif
                                                </td>
                                                <td class="dien_thoai">{{ $data->dien_thoai }}</td>
                                                <td class="d-none cccd">{{ $data->cccd }}</td>
                                                <td class="comment">{{ $data->comment }}</td>
                                                <td class="menh_gia">{{ is_numeric(@$data->menh_gia) ? number_format(@$data->menh_gia + 50000000) : @$data->menh_gia }}</td>
                                                <td style="width:200px;" class="type_call">
                                                    @foreach (\App\Models\Customer::getInforOption() as $key => $value)
                                                        @if ($key == $data->type_call)
                                                            <span data-id="{{ $key }}" class="badge {{ $key == 0 ? 'bg-danger' : 'bg-primary' }} ">{{ $value }}</span>
                                                        @endif
                                                    @endforeach
                                                </td>
                                                <td class="dia_chi_cu_the">{{ $data->dia_chi_cu_the }}</td>
                                                <td class="updated_at">{{ $data->updated_at }}</td>
                                                <td class="d-none area_name">{{ $data->name }}</td>
                                            </tr>
                                            @php
                                            $j--;
                                            @endphp
                                            @endforeach
                                        </tbody>
                                    </table>
                                    {!! $history->histories->links('_partials.pagination') !!}
                                    <!-- End Table with hoverable rows -->
                                </div>
                            </div>
                        </div>
                    </div>
                </div>
            </div>
            <button type="button" class="btn btn-primary d-none" id="ratings" data-bs-toggle="modal" data-bs-target="#largeModal"></button>

            <div class="modal fade" id="largeModal" tabindex="-1" aria-hidden="true">
                <div class="modal-dialog">
                  <div class="modal-content">
                    <div class="modal-header">
                      <h5 class="modal-title"><b class="text-danger">BẢNG XẾP HẠNG NHÂN VIÊN GỌI KHÁCH HÀNG TRONG NGÀY HÔM NAY.</b></h5>
                      <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Close"></button>
                    </div>
                    <div class="modal-body">
                        <table class="table is-ratings">
                            <thead>
                            <tr>
                                <th scope="col">Xếp Hạng</th>
                                <th scope="col">Tên Nhân Viên</th>
                            </tr>
                            </thead>
                            <tbody></tbody>
                        </table>
                    </div>
                    <div class="modal-footer justify-content-between">
                        <p class="text-left text-success">CẢM ƠN SỰ CỐNG HIẾN VÀ LÀM VIỆC NGHIÊM TÚC CỦA CÁC BẠN TRONG NGÀY HÔM NAY.</p>
                      <button type="button" class="btn btn-secondary w-100" data-bs-dismiss="modal">Đóng</button>
                    </div>
                  </div>
                </div>
        </div>
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

    <script>
        $(document).ready(function() {
            var formData = new FormData();
            formData.append('_token', "{{ csrf_token() }}");

            setInterval(function() {
                var d = new Date();
                var t = d.toLocaleTimeString();

                if (t == '17:20:00' || t == '5:20:00 PM') {

                    $('body').append(`<div id="film-container">
                        <div class="animated-flicker">
                            <svg id="film-mask" version="1.1" xmlns="http://www.w3.org/2000/svg"
                                xmlns:xlink="http://www.w3.org/1999/xlink" xml:space="preserve"
                                width="100%" height="100%" preserveAspectRatio="none" viewBox="0 0 400 225">
                                <rect class="rectangle01" width="400" height="225"/>
                                <line id="line-h" class="line01" x1="0" y1="112.5" x2="400" y2="112.5"/>
                                <line id="line-v" class="line02" x1="200" y1="0" x2="200" y2="225"/>
                                </svg>
                                <svg id="film-countdown" version="1.1" xmlns="http://www.w3.org/2000/svg"
                                xmlns:xlink="http://www.w3.org/1999/xlink" xml:space="preserve"
                                width="100%" height="100%" viewBox="0 0 400 225">
                                <circle id="circle-outer" class="circle01" cx="200" cy="112.5" r="95"/>
                                <circle id="circle-inner" class="circle01" cx="200" cy="112.5" r="85"/>
                                <circle class="circle02 animated-rotate" cx="200" cy="112.5" r="494.5"/>
                                    <g id="numbers" text-anchor="middle" class="no-select">
                                        <text id="animated-text1" x="200.5" y="155">10</text>
                                        <text id="animated-text2" x="200.5" y="155">9</text>
                                        <text id="animated-text3" x="200.5" y="155">8</text>
                                        <text id="animated-text4" x="200.5" y="155">7</text>
                                        <text id="animated-text5" x="200.5" y="155">6</text>
                                        <text id="animated-text6" x="200.5" y="155">5</text>
                                        <text id="animated-text7" x="200.5" y="155">4</text>
                                        <text id="animated-text8" x="200.5" y="155">3</text>
                                        <text id="animated-text9" x="200.5" y="155">2</text>
                                        <text id="animated-text10" x="200.5" y="155">1</text>
                                    </g>
                                </svg>
                            </div>
                        </div>
                    `);
                    setTimeout(() => {
                        $('#film-container').addClass('d-none');
                        if ($(".is-ratings tbody tr").length > 0) {
                            $("#ratings").trigger('click');
                        }
                    }, 10500);

                    $.ajax({
                        url: "{{route('ratings')}}",
                        data: formData,
                        type: 'post',
                        async: false,
                        processData: false,
                        contentType: false,
                        success:function(response) {
                            if(response.data.length > 0) {
                                response.data.forEach(function(item, index) {
                                    let class_color = (index == 0) ? 'danger' : (index == 1 ? 'success' : (index == 2 ? 'primary' : ''));

                                    $(".is-ratings tbody").append(`
                                        <tr>
                                            <th scope="row">
                                                <span class="${index <= 2 ? 'badge' : '' }  rounded-pill bg-${class_color}"> ${index + 1}</span>
                                            </th>
                                            <td>${item.name}</td>
                                        </tr>
                                    `);
                                });
                            }
                        },
                        error: function(errors) {
                            console.log(errors);
                        }
                    });
                }
            }, 1000)

            document.querySelector(".is-copy").onclick = (e) => {
                navigator.clipboard.writeText($(e.currentTarget)[0].outerText);
                if($(e.currentTarget)[0].outerText) {
                    if($('.is-status-clip').hasClass('bi bi-clipboard')) {
                        $('.is-status-clip').removeClass('bi bi-clipboard');
                        $('.is-status-clip').addClass('bi bi-clipboard-check');
                    }
                    setTimeout(() => {
                        $('.is-status-clip').removeClass('bi bi-clipboard-check');
                        $('.is-status-clip').addClass('bi bi-clipboard');
                    }, 2000);
                }
            }

            $('.btn-search').click(function() {
                var area_id = $("#data_area_id").val();
                $('#form-search').append(`<input type="hidden" name="area_id" value="${area_id}" />`);
                $( "#form-search" ).submit();
            });

            if (!$('#id_contract').val()) {
                toastr.options.progressBar = true;
                toastr.options.closeButton = true;
                toastr.info('Vui lòng chuyển đổi nguồn dữ liệu hoặc liên hệ với quản trị viên để cấp dữ liệu thêm.', 'Thông Báo');
            }
            $(".is-item-customer").click(function() {
                $(".btn-save").removeClass('d-none');

                var el= $(this).get(0);

                $('#customer_id').val($($(el).find('.id').get(0)).text());
                $('#ngay_tham_gia').val($($(el).find('.ngay_tham_gia').get(0)).text());
                $('#nam_dao_han').val($($(el).find('.nam_dao_han').get(0)).text());
                $('#id_contract').val($($(el).find('.so_hop_dong').get(0)).text());
                $('#vpbank').val($($(el).find('.vpbank').get(0)).text());
                $('#cv').val($($(el).find('.cv').get(0)).text());
                $('#menh_gia').val($($(el).find('.menh_gia').get(0)).text());
                $('#ten_kh').val($($(el).find('.ten_kh').get(0)).text().trim());
                $('#dien_thoai').val($($(el).find('.dien_thoai').get(0)).text());
                $('#cccd').val($($(el).find('.cccd').get(0)).text());
                $('#dia_chi_cu_the').val($($(el).find('.dia_chi_cu_the').get(0)).text());
                $('#tuoi').val($($(el).find('.tuoi').get(0)).text().trim());
                $('#data_area_id').val($($(el).find('.area_id').get(0)).text().trim());
                $('#comment').val($($(el).find('.comment').get(0)).text());
                var id_result = $($(el).find('.type_call')).children().data('id');

                var sex = $($(el).find('.gioi_tinh').get(0)).text().trim();
                $("#gioi_tinh option").each(function() {
                    $(this).attr('selected', false);
                    if ($(this).text().trim() == sex) {
                        $(this).attr('selected', true);
                    } else if($(this).text().trim() == sex) {
                        $(this).attr('selected', true);
                    }
                });

                var area_text = $($(el).find('.area_name').get(0)).text().trim();

                $("#data_area_id option").each(function() {
                    if ($(this).text().trim() == area_text) {
                        $(this).prop("selected","selected");
                    }
                });

                $('.is-result .list-group-item').each(function() {
                    var value_radio = $($(this).find('input').get(0)).val();
                    if (value_radio  == id_result) {
                        $("#gridRadios"+ id_result).trigger('click');
                    }
                });
            });
        });
    </script>

    {!! Toastr::message() !!}

</body>

</html>
