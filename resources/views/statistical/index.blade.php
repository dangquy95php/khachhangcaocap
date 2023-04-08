@section('title','Thống kê')
@extends('layouts.template')

@section('breadcrumb')

   <h1>THỐNG KÊ</h1>

    {{ Breadcrumbs::render('statistical') }}

@endsection

@section('content')

<section class="section">

   <div class="row">
      <div class="col-lg-7">
         <div class="card">
            <div class="card-body pt-2">
               <!-- Table with stripped rows -->
               <table class="table">
                  <thead>
                     <tr>
                        <th scope="col">#</th>
                        <th scope="col">Tên Nhân Viên</th>
                        <th scope="col">Trạng Thái</th>
                        <th scope="col">Ghi Chú</th>
                        <th scope="col">Ngày Tạo</th>
                        <th scope="col" class="text-center">
                            @if(\Request::route()->getName() == 'edit_area')
                                <a class="btn btn-success" href="{{route('index_area')}}">Tạo mới</a>
                            @else
                                <span>Hành Động</span>
                            @endif
                        </th>
                     </tr>
                  </thead>
                  <tbody>
                    @foreach($areas as $item)
                    <tr>
                        <td>{{ $item->id }}</td>
                        <td>{{ $item->name }}</td>
                        <td>{!! $item->status == 1 ? '<span class="badge bg-success">Đang Mở</span>' : '<span class="badge bg-secondary">Chưa kích hoạt</span>' !!}</td>
                        <td>{{ $item->note }}</td>
                        <td>{{ date('d-m-Y', strtotime($item->created_at)) }}</td>
                        <td class="text-center">
                            <a data-bs-target="#deleteModal{{ $item->id }}" data-bs-toggle="modal" class="btn btn-danger">Xoá</a>

                            <a href="{{route('edit_area', $item->id)}}" class="btn btn-primary">Sửa</a>
                        </td>
                    </tr>

                    <div class="modal fade" id="deleteModal{{ $item->id }}" tabindex="-1" aria-hidden="true">
                        <div class="modal-dialog">
                            <div class="modal-content">
                            <div class="modal-header">
                                <h5 class="modal-title">Bạn Có Muốn Xoá Không?</h5>
                                <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Close"></button>
                            </div>
                            <div class="modal-body">
                                <p>Tên Khu Vực: <b>{{ $item->name }}</b></p>
                                <p>Trạng Thái: {!! $item->status == 1 ? '<span class="badge bg-success">Đang Mở</span>' : '<span class="badge bg-secondary">Chưa kích hoạt</span>' !!}</p>
                                <p>Ghi Chú: {{ $item->note }}</p>
                            </div>
                            <div class="modal-footer">
                                <button type="button" class="btn btn-secondary" data-bs-dismiss="modal">Đóng</button>
                                <a type="button" href="{{route('delete_area', $item->id)}}" class="btn btn-primary">Đồng Ý</a>
                            </div>
                            </div>
                        </div>
                    </div>
                    @endforeach
                  </tbody>
               </table>
               <!-- End Table with stripped rows -->
            </div>
         </div>
      </div>
      <div class="col-lg-5">
        <div class="card">
            <div class="card-body">
            <h5 class="card-title text-center mb-0"><b>{{ \Request::route()->getName() == 'index_area' ? 'THÊM KHU VỰC' : 'CHỈNH SỬA KHU VỰC' }}</b></h5>
                <!-- Vertical Form -->
                <form class="row g-3" method="post" action="{{ \Request::route()->getName() == 'index_area' ? route('create_area') : '' }}">
                    @csrf
                    <div class="col-12">
                        <label for="inputNanme4" class="form-label">Tên Khu Vực</label>
                        <input type="text" value="{{ @$area->name }}" name="name" class="form-control" id="inputNanme4">
                        @include('_partials.alert', ['field' => 'name'])
                    </div>
                    <div class="col-12">
                        <label for="inputEmail4" class="form-label">Trạng Thái</label>
                        <select class="form-select" name="status" aria-label="Default select example">
                            <option selected="0">Vui lòng chọn trạng thái</option>
                            @foreach($area_status as $item)
                                <option {{ @$area->status == $item ? 'selected' : ''; }} value="{{ $item }}">{{ $item == 0 ? 'Chưa kích hoạt' : 'Đang mở'; }}</option>
                            @endforeach
                        </select>
                    </div>
                    <div class="col-12">
                        <label for="inputNanme5" class="form-label">Ghi chú</label>
                        <textarea type="text" placeholder="Ghi chú" rows="8" name="note" class="form-control" id="inputNanme5">{{ @$area->note }}</textarea>
                    </div>
                    <div class="text-center">
                        <button type="submit" class="btn btn-primary">{{ \Request::route()->getName() == 'index_area' ? 'Tạo Mới' : 'Cập Nhật' }}</button>
                    </div>
                </form><!-- Vertical Form -->
            </div>
        </div>
      </div>
   </div>
</section>
@endsection
