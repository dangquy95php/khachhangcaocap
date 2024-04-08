@section('title','Trang chủ')
@extends('layouts.template')

@section('breadcrumb')

   <h1>DANH SÁCH KHU VỰC</h1>

    {{ Breadcrumbs::render('area') }}

@endsection

@section('content')

<section class="section">

   <div class="row">
      <div class="col-lg-8 pe-2">
         <div class="card">
            <div class="card-body pt-2 table-responsive">
               <!-- Table with stripped rows -->
               <table class="table datatable" style="min-width:695px;">
                  <thead>
                     <tr>
                        <th scope="col">#</th>
                        <th style="width:210px;" scope="col">Tên Nhân Viên</th>
                        <th scope="col">Trạng Thái</th>
                        {{-- <th scope="col">Ghi Chú</th> --}}
                        <th scope="col">Ngày Tạo</th>
                        <th scope="col" class="text-center" style="min-width:250px;">
                            @if(\Request::route()->getName() == 'edit_area')
                                <a class="btn btn-success text-white" href="{{route('index_area')}}">Tạo mới</a>
                                @if(count($areas) > 0)
                                <a class="btn btn-danger text-white" onclick="return confirm('Bạn có muốn tất cả các khu vực không?');" href="{{route('delete_area_all')}}">Xoá Tất Cả</a>
                                @endif
                            @else
                                <span>Hành Động</span>
                                @if(count($areas) > 0)
                                <a class="btn btn-danger text-white" onclick="return confirm('Bạn có muốn tất cả các khu vực không?');"  href="{{route('delete_area_all')}}">Xoá Tất Cả</a>
                                @endif
                            @endif
                        </th>
                     </tr>
                  </thead>
                  <tbody>
                    @php
                    $i = count($areas);
                    @endphp

                    @foreach($areas as $item)
                    <tr>
                        <td>{{ $i }}</td>
                        <td>
                            <button type="button" class="btn btn-primary pe-1">
                                {{ $item->name }} <span class="ms-2 badge bg-{{ count($item->customers) > 0 ? 'white' : 'danger' }} text-{{ count($item->customers) > 0 ? 'primary' : 'white' }}">{{count($item->customers)}}</span>
                            </button>
                        </td>
                        <td>
                            <span class="badge bg-{{$item->status == 1 ? 'success' : 'secondary' }}">{{ $item->status == 1 ? 'Đang Mở' : 'Chưa kích hoạt' }}</span>
                        </td>
                        {{-- <td>{{ $item->note }}</td> --}}
                        <td>{{ date('d-m-Y', strtotime($item->updated_at)) }}</td>
                        <td class="text-center">
                            <a onclick="return confirm(`Bạn có muốn xoá khu vực {{ $item->name }} không?`);" href="{{route('delete_area', $item->id)}}"  class="btn btn-danger">Xoá</a>
                            <a href="{{route('edit_area', $item->id)}}" class="btn btn-primary">Sửa</a>
                            <a onclick="return confirm('Bạn có muốn khôi phục lại dữ liệu cho khu vực {{ $item->name }} không?');" href="{{route('reopen_area', $item->id)}}" class="text-white btn ps-1 {{empty($item->user_id) ? 'btn-dark disabled': 'btn-warning'}}"><i class="bi bi-back"></i>{{empty($item->user_id) ? 'Không Thể': 'Khôi Phục'}}</a>
                        </td>
                    </tr>
                    @php
                    $i--;
                    @endphp
                    @endforeach
                  </tbody>
               </table>
               <!-- End Table with stripped rows -->
               {!! $areas->links('_partials.pagination') !!} 
            </div>
         </div>
      </div>
      <div class="col-lg-4 ps-0">
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
                            @foreach($areaAtatus as $item)
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
