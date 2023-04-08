@section('title','Trang chủ')
@extends('layouts.template')

@section('breadcrumb')

   <h1>DANH SÁCH NGƯỜI DÙNG</h1>

   {{ Breadcrumbs::render('list_account') }}

@endsection

@section('content')

<section class="section">
   <div class="row">
      <div class="col-lg-12">
         <div class="card">
            <div class="card-body table-responsive">
               <!-- Table with stripped rows -->
               <table class="table table-striped">
                  <thead>
                     <tr>
                        <th scope="col">#</th>
                        <th scope="col">Tên Nhân Viên</th>
                        <th scope="col">Tên Đăng Nhập</th>
                        <th scope="col">Quyền Truy Cập</th>
                        <th scope="col">Ngày Tạo</th>
                        <th scope="col">Trạng Thái</th>
                        <th scope="col" class="text-center">
                            <a href="{{ route('create_account') }}" type="button" class="btn btn-primary">Thêm</a>
                        </th>
                     </tr>
                  </thead>
                  <tbody>
                     @foreach($data as $key => $item)
                     <tr class="{{$item->username == \Auth::user()->username ? 'bg-danger text-white' : '' }}">
                        <th scope="row">{{ ++$key }}</th>
                        <td>{{$item->name}}</td>
                        <td>{{$item->username}}</td>
                        <td>
                            @if($item->role == '1')
                                <span class="badge bg-primary">Người dùng</span>
                            @endif
                            @if($item->role == '2')
                                <span class="badge bg-success">Quản trị</span>
                            @endif
                        </td>
                        <td>{{ $item->created_at }}</td>
                        <td>
                           @if ($item->status == 1)
                              <span class="badge rounded-pill bg-success">Đang hoạt động</span>
                           @endif

                           @if($item->status == 2)
                              <span class="badge rounded-pill bg-danger">Không hoạt động</span>
                           @endif

                           @if($item->status == 0)
                              <span class="badge rounded-pill bg-secondary">Chưa được duyệt</span>
                           @endif
                        </td>
                        <td class="text-center">
                           <a class="btn btn-warning text-white" href="{{route('edit_account', $item->id) }}">Sửa</a>
                           @if($item->username != \Auth::user()->username)
                              <a onclick="return confirm(`Bạn có muốn xóa nhân viên {{ $item->username }} không?`);"  href="{{route('delete_account', $item->id) }}" class="btn btn-danger">Xoá</button>
                            @endif
                        </td>
                     </tr>
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
