@section('title','Trang chủ')
@extends('layouts.template')

@section('breadcrumb')

   <h1>DANH SÁCH EXCEL IMPORT NGƯỜI DÙNG</h1>

    {{ Breadcrumbs::render('excel') }}

@endsection

@section('content')

<section class="section">

    <div class="row">
        <div class="col-lg-12">
            <div class="card">
                <div class="card-body pt-3">
                    <div id="overlay">
                        <div class="cv-spinner">
                            <span class="spinner"></span>
                        </div>
                    </div>
                    <!-- General Form Elements -->
                    <form id="form_import_excel" method="POST" enctype="multipart/form-data">
                        @csrf
                        <div class="row mb-3">
                            <div class="col-12">
                                <label for="inputText" class="col-sm-2 col-form-label"><b>Chọn file Excel Import:</b></label>
                                <span class="badge rounded-pill bg-warning instro" role="button" data-bs-toggle="modal" data-bs-target="#fullscreenModal">Hướng dẫn sử dụng</span>
                                <div class="is-show-error"></div>

                                <input class="form-control" name="file" type="file" id="formFile">
                            </div>
                        </div>
                        <a class="btn-import btn btn-primary me-2">Improt File</a>
                        <!-- <a href="{{route('export_customer')}}" class="btn btn-success">Export File</a> -->
                    </form>

                    <div class="modal fade" id="fullscreenModal" tabindex="-1">
                        <div class="modal-dialog modal-fullscreen">
                          <div class="modal-content">
                            <div class="modal-header">
                              <h5 class="modal-title"><b>Hướng Dẫn import Excel</b></h5>
                              <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Close"></button>
                            </div>
                            <div class="modal-body">
                              <p class="mb-1"><b>Trường số hợp đồng là trường bắt buộc phải có</b></p>
                              <p class="mb-1"><b>Điền thêm một số trường giống hình mẫu excel sẽ đầy đủ dữ liệu để nhân viên gọi dễ dàng hơn.</b></p>
                              <p for=""><b>Dòng đầu tiên không được để trống</b></p>
                              <img style="width:100%;" src="{{ asset('assets/img/intro.png') }}" alt="">
                            </div>
                          </div>
                        </div>
                    </div><!-- End Full Screen Modal-->

                    <!-- End General Form Elements -->
                </div>
            </div>
        </div>
    </div>

   <div class="row">
      <div class="col-lg-12">
         <div class="card">
            <div class="card-body table-responsive">
               <!-- Table with stripped rows -->
               <div class="d-md-flex justify-content-between card-title pb-0 pt-2">
                   <h5 class="mb-0 d-flex align-items-center">LỊCH SỬ IMPORTED DỮ LIỆU</h5>

                   <form class="row g-3" method="GET" enctype="multipart/form-data" action="/admin/excel/search">
                        <div class="col-auto">
                            <input name="search" type="text" class="form-control" placeholder="Tìm kiếm số hợp đồng ">
                        </div>
                        <div class="col-auto">
                            <button type="submit" class="btn btn-primary">Tìm kiếm</button>
                        </div>
                    </form>
               </div>
               <hr class="mt-0">
               <table class="table table-striped">
                  <thead>
                     <tr>
                        <th scope="col">#</th>
                        <th scope="col">Người Import</th>
                        <th scope="col">Số dòng import thành công</th>
                        <th scope="col">Thông tin</th>
                        <th scope="col"><i class="bi bi-file-earmark-excel text-success"></i>Tên File</th>
                        <th scope="col">Bắt đầu import</th>
                        <th scope="col">Kết thúc import</th>
                        <!-- <th scope="col" class="text-center">Hành Động</th> -->
                     </tr>
                  </thead>
                  <tbody>
                    @php
                    $j = $importHistory->total();
                    if ($importHistory->currentPage() >= 2) {
                       $j = $importHistory->total() - (($importHistory->currentPage() - 1) * $importHistory->perPage());
                    }
                    @endphp
                    @foreach($importHistory as $import)
                     <tr>
                        <th scope="row">{{ $j }}</th>
                        <td>{{ $import->user->name }}</td>
                        <td><span class="badge rounded-pill {{ $import->number != 0 ? 'bg-success' : 'bg-warning' }} ">{{ $import->number }}</span></td>
                        <td><span class="badge rounded-pill {{ $import->status == 'Thành Công' ? 'bg-success' : ($import->status == 'Trùng Lặp' ? 'bg-warning' : 'bg-danger' ) }} ">{{ $import->status }}</span></td>
                        <td><i class="bi bi-file-earmark-excel text-success"></i>{{ $import->file_name }}</td>
                        <td>{{ $import->created_at }}</td>
                        <td>{{ $import->updated_at }}</td>
                        <!-- <td class="text-center">
                            <a href="{{route('delete_import', $import->id) }}" class="btn btn-danger">Xoá</button>
                        </td> -->
                     </tr>
                     @php
                    $j--;
                    @endphp
                    @endforeach
                  </tbody>
               </table>
               <!-- End Table with stripped rows -->
               {!! $importHistory->links('_partials.pagination') !!}
            </div>
         </div>
      </div>
   </div>
</section>
@endsection

@push('scripts')
<script type="text/javascript">

    $(document).ready(function() {
        $(".instro").click(function() {
            $(this).trigger('click');
        });

        var formImport = $('#form_import_excel');
        $(".btn-import").click(function(event){
            event.preventDefault();
            $("#overlay").show();

            toastr.options.positionClass = 'toast-top-full-width';
            toastr.warning('Dữ liệu đang chạy. Không nên tắt.','',{timeOut: 0})
            var formData = new FormData();
            if ($('#formFile').get(0).files.length !== 0) {
                formData.append('file', $('#formFile')[0].files[0]);
            }
            formData.append('_token', "{{ csrf_token() }}");
            toastr.options.positionClass = 'toast-top-right';
            $.ajax({
                url: "{{route('data_import')}}",
                data: formData,
                type: 'post',
                async: true,
                processData: false,
                contentType: false,
                success:function(response) {
                    $("#overlay").hide();
                    location.reload();
                },
                error: function(errors) {
                    toastr.remove();
                    toastr.error('Import dữ liệu thất bại!')
                    $("#overlay").hide();
                    if (errors.status == 422) {
                        var errorsData = errors.responseJSON.errors.file;

                        $.each(errorsData , function (index, value){
                            $(".is-show-error").html(
                                `<div class="alert alert-danger alert-dismissible fade show d-block" role="alert">
                                ${value}
                                <button type="button" class="btn-close" data-bs-dismiss="alert" aria-label="Close"></button>
                            </div>`
                            );
                        });
                    }
                    if (errors.status == 400 || errors.status == 500) {
                        var error = errors.responseJSON.message;
                        $(".is-show-error").html(
                            `<div class="alert alert-danger alert-dismissible fade show d-block" role="alert">
                                ${error}
                                <button type="button" class="btn-close" data-bs-dismiss="alert" aria-label="Close"></button>
                            </div>`
                        );
                    }
                }
            });
        });
    })
</script>
@endpush
