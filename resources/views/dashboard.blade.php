@section('title','Trang chủ')
@extends('layouts.template')

@section('breadcrumb')

   <h1>TỔNG QUAN DASHBOARD</h1>

   {{ Breadcrumbs::render('home') }}

@endsection

@section('content')

<section class="section dashboard">
   <div class="row">
      <!-- Left side columns -->
      <div class="col-lg-12">
         <div class="row">
            <!-- Customers Card -->
            <!-- <div class="col-xxl-4 col-md-6">
               <div class="card info-card customers-card">
                  <div class="card-body">
                     <h5 class="card-title">TỔNG KHÁCH HÀNG IMPORT <span>| Hôm Nay</span></h5>
                     <div class="d-flex align-items-center">
                        <div class="card-icon rounded-circle d-flex align-items-center justify-content-center">
                           <i class="bi bi-people"></i>
                        </div>
                     </div>
                  </div>
               </div>
            </div> -->
            <!-- End Customers Card -->
            <!-- Sales Card -->
            <div class="col-xxl-6 col-md-6">
               <div class="card info-card sales-card">
                  <div class="card-body">
                     <h5 class="card-title">TỔNG KHÁCH ĐÃ GỌI <span>| Hôm Nay</span></h5>
                     <div class="d-flex align-items-center">
                        <div class="card-icon rounded-circle d-flex align-items-center justify-content-center">
                           <i class="bi bi-headset"></i>
                        </div>
                        <div class="ps-3">
                           <h6>{{ $totalCallCustomer }}</h6>
                        </div>
                     </div>
                  </div>
               </div>
            </div>
         </div>
      </div>
      <!-- Recent Sales -->
      <div class="col-12">
         <div class="card recent-sales overflow-auto">
            <div class="card-body table-responsive">
                <div class="d-flex align-items-center">
                    <h5 class="card-title">Khách Hàng Đã Hẹn <span>| Hôm nay</span></h5>
                    <a href="{{route('appointment_excel')}}" class="btn btn-outline-success pt-1 pb-1 ms-3 ps-1">
                     <i class="bi bi-file-earmark-excel"></i> Xuất Excel</a>
                </div>
               <table class="table datatable">
                  <thead>
                     <tr>
                        <th scope="col">#</th>
                        <th scope="col">Số Hợp Đồng</th>
                        <th style="min-width: 150px;" scope="col">Tên Khách Hàng</th>
                        <th scope="col">Số Điện Thoại</th>
                        <th style="min-width: 200px;" scope="col">Ghi Chú</th>
                        <th scope="col">Tài Khoản Gọi</th>
                        <th scope="col">Giới Tính</th>
                        <th scope="col">Tuổi</th>
                        <th style="min-width: 250px;" scope="col">Địa Chỉ</th>
                        <th style="min-width: 100px;" scope="col">Thời Gian Gọi</th>
                     </tr>
                  </thead>
                  <tbody>
                     @php
                     $i = count($result);
                     @endphp
                     @foreach($result as $data)
                        <tr>
                           <th scope="row">{{$i}}</th>
                           <td>{{ $data->so_hop_dong }}</td>
                           <td>{{ $data->ten_kh }}</td>
                           <td>{{ $data->dien_thoai }}</td>
                           <td>{{ $data->comment }}</td>
                           <td>{{ $data->username }}</td>
                           <td>
                              @if($data->gioi_tinh == 'M')
                              Nam
                              @endif
                              @if($data->gioi_tinh == 'F')
                              Nữ
                              @endif
                           </td>
                           <td>{{ $data->tuoi }}</td>
                           <td>{{ $data->dia_chi_cu_the }}</td>
                           <td>{{ $data->updated_at }}</td>
                        </tr>
                        @php
                        $i--;
                        @endphp
                     @endforeach
                  </tbody>
               </table>
            </div>
         </div>
      </div>
      <!-- End Left side columns -->
   </div>
</section>

@endsection
