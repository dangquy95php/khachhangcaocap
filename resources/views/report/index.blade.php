@section('title','Báo cáo cuộc gọi khách hàng')
@extends('layouts.template')

@section('breadcrumb')

   <h1>BÁO CÁO CUỘC GỌI KHÁCH HÀNG HÔM NAY</h1>

   {{ Breadcrumbs::render('report') }}

@endsection

@section('content')

<section class="section">
   <div class="row">
      <div class="col-lg-8">
         <div class="card">
            <div class="card-body table-responsive pt-3">
               <!-- Table with stripped rows -->
               <table data-toggle="table" class="table table-striped datatable">
                  <thead>
                     <tr>
                     <th scope="col">#</th>
                     <th scope="col">Nhân Viên</th>
                     <th class="text-center" scope="col">Tổng Cuộc Gọi</th>
                     <th class="text-center" scope="col">Tổng Cuộc Gọi Đã Hẹn</th>
                     </tr>
                  </thead>
                  <tbody>
                    @php
                    $i = count($listCallOfStaff);
                    @endphp
                    @foreach($listCallOfStaff as $user)
                        @if($user->username != 'admin')
                            <tr>
                                <th scope="row">{{ $i }}</th>
                            <td>{{ $user->name ?: $user->username }}</td>
                            <td class="text-center">
                                <span class="badge bg-{{ count($user->customers_today_called) > 0 ? 'success' : 'secondary' }}">{{ count($user->customers_today_called) }}</span>
                                </td>
                            <td class="text-center">
                                    <span class="badge bg-{{ $user->customers_today_called->where('type_call', 0)->count() > 0 ? 'primary' : 'secondary' }}">{{ $user->customers_today_called->where('type_call', 0)->count() }}</span>

                                </td>
                            </tr>
                        @endif
                    @php
                    $i--;
                    @endphp
                    @endforeach
                  </tbody>
               </table>
               <!-- End Table with stripped rows -->
            </div>
         </div>
      </div>
   </div>
</section>

@endsection
