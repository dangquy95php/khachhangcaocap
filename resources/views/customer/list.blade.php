@section('title','Danh sách khách hàng')
@extends('layouts.template')

@section('breadcrumb')

   <h1>DANH SÁCH KHÁCH HÀNG IMPORT</h1>

   {{ Breadcrumbs::render('list_customer') }}

@endsection

@section('content')

<section class="section">
   <div class="row">
      <div class="col-lg-12">
         <div class="card">
            <form class="pt-3" method="GET" action="{{ route('customer.delete') }}">
               @csrf
               <div class="row align-items-center">
                  <div class="col-md-4 col-sm-5">
                     <div class="row p-2">
                        <label for="inputTime" class="col-md-5 col-form-label text-md-end text-sm-start"><b>Ngày bắt đầu:</b></label>
                        <div class="col-md-7">
                           <input type="date" name="start_date" class="form-control">
                           @if($errors->has('start_date'))
                                <p class="text-danger">{{ $errors->first('start_date') }}</p>
                            @endif
                        </div>
                     </div>
                  </div>
                  <div class="col-md-4 col-sm-5">
                     <div class="row p-2">
                        <label for="inputTime" class="col-md-5 col-form-label text-md-end text-sm-start"><b>Ngày kết thúc:</b></label>
                        <div class="col-md-7">
                           <input type="date" name="end_date" class="form-control">
                            @if($errors->has('end_date'))
                                <p class="text-danger">{{ $errors->first('end_date') }}</p>
                            @endif
                        </div>
                     </div>
                  </div>
                  <div class="col-md-4 col-sm-2">
                     <div class="px-md-2 p-sm-0">
                        <button value="delete" onclick="return confirm('Dữ liệu xóa sẽ không khôi phục được. Bạn có muốn xóa không?');" name="action" type="submit" class="btn btn-primary"><i class="bi bi-trash"></i> Xoá</button>
                        @if(\Auth::user()->username == 'PHANYEN')
                           <button name="action" value="export" class="btn btn-success" type="submit"><i class="bi bi-file-earmark-excel text-white"></i> Export</button>
                        @endif
                     </div>
                  </div>
               </div>
            </form>
            <hr/>
            <div class="card-body table-responsive">
               <!-- Table with stripped rows -->
               <table class="table table-striped">
                  <thead>
                     <tr>
                     <th scope="col">#</th>
                     <th scope="col">Số hợp đồng</th>
                     <th scope="col">Ngày tham gia</th>
                     <th scope="col">Mệnh Giá</th>
                     <th scope="col">Năm Đáo Hạn</th>
                     <th scope="col">Họ Và Tên</th>
                     <th scope="col">Giới Tính</th>
                     <th scope="col">Ngày Sinh</th>
                     <th scope="col">Điện Thoại</th>
                     <th scope="col">Tuổi</th>
                     <th scope="col text-center">Địa chỉ</th>
                     <th scope="col text-center">Ngày Tạo</th>
                     <th scope="col"></th>
                     </tr>
                  </thead>
                  <tbody>
                     @php
                     $j = $customers->total();
                     if ($customers->currentPage() >= 2) {
                        $j = $customers->total() - (($customers->currentPage() - 1) * $customers->perPage());
                     }
                     @endphp
                    @foreach($customers as $customer)
                     <tr>
                        <th scope="row">{{ $j }}</th>
                           <td>{{ $customer->so_hop_dong }}</td>
                           <td>{{ $customer->ngay_tham_gia }}</td>
                           <td>{{ is_numeric(@$customer->menh_gia) ? number_format(@$customer->menh_gia) : @$customer->menh_gia }}</td>
                           <td>{{ $customer->nam_dao_han }}</td>
                           <td>{{ $customer->ten_kh }}</td>
                           <td>
                              @if($customer->gioi_tinh == 'M')
                              Nam
                              @endif
                              @if($customer->gioi_tinh == 'F')
                              Nữ
                              @endif
                           </td>
                           <td>{{ $customer->ngay_sinh }}</td>
                           <td>{{ $customer->dien_thoai }}</td>
                           <td>{{ $customer->tuoi }}</td>
                           <td>{{ $customer->dia_chi_cu_the }}</td>
                           <td>{{ $customer->created_at }}</td>
                           <td>
                              <a class="btn btn-danger d-flex align-items-center px-2" onclick="return confirm('Dữ liệu xóa sẽ không khôi phục được. Bạn có muốn xóa không?');" href="{{ route('customer.delete.byId', $customer->id) }}"><i class="bi bi-trash"></i>&nbsp;Xoá</a>
                           </td>
                        </tr>
                        @php
                        $j--;
                        @endphp
                     @endforeach
                  </tbody>
               </table>
               <!-- End Table with stripped rows -->
               {!! $customers->links('_partials.pagination') !!}
            </div>
         </div>
      </div>
   </div>
</section>

@endsection
