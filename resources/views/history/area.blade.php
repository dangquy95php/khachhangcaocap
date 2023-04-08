@section('title','Lịch sử khôi phục dữ liệu')
@extends('layouts.template')

@section('breadcrumb')

   <h1>LỊCH SỬ KHÔI PHỤC DỮ LIỆU</h1>

   {{ Breadcrumbs::render('history_area') }}

@endsection

@section('content')

<section class="section">
    <div class="row">
        <div class="col-lg-9 col-md-12">
            <div class="card overflow-auto">
                <div class="card-body table-responsive-md">
                    <h5 class="card-title">Qúa trình khôi phục dữ liệu của khu vực.</h5>
                    <!-- Small tables -->
                    <table class="table datatable">
                    <thead>
                        <tr>
                        <th scope="col">#</th>
                        <th style="min-width:115px;" scope="col">Tên khu vực</th>
                        <th scope="col">Người khôi phục</th>
                        <th scope="col">Người đã gọi</th>
                        <th scope="col">Ngày khôi phục</th>
                        </tr>
                    </thead>
                    <tbody>
                        @php
                        $i = count($data);
                        @endphp
                        @foreach($data as $area)
                        <tr>
                            <th scope="row">{{ $i }}</th>
                            <td>
                                <button type="button" class="btn btn-primary mb-2 ps-1 pe-1">
                                    {{ $area->area->name }} <span class="badge bg-white text-primary ms-2">{{$area->count_record}}</span>
                              </button>
                            </td>
                            <td class="text-center"><span class="badge bg-success">{{$area->author->username}}</span></td>
                            <td class="text-center"><span class="badge bg-danger">{{$area->user->username ?? 'Rỗng'}}</span></td>
                            <td>{{ $area->created_at }}</td>
                        </tr>
                        @php
                        $i--;
                        @endphp
                        @endforeach
                    </tbody>
                    </table>
                    <!-- End small tables -->
                    {!! count($data) == 0 ? '<h5 class="text-center pt-5 pb-5"><b>KHÔNG CÓ LỊCH SỬ KHÔI PHỤC</b></h5>' : '' !!}
                </div>
            </div>
        </div>
    </div>
</section>
@endsection
