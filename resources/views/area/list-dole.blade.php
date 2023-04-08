@section('title','Trang chủ')
@extends('layouts.template')

@section('breadcrumb')

   <h1>QUẢNG LÝ KHÁCH HÀNG THEO TỪNG KHU VỰC</h1>

    {{ Breadcrumbs::render('area.dole') }}

@endsection

@section('content')

<section class="section">
    <form action="" method="post" id="form-customer">
        @csrf
            <div class="row">
                <div class="card mb-3">
                    <div class="card-body pb-2">
                      <div class="accordion accordion-flush" id="accordionFlushExample">
                        <div class="accordion-item">
                            <h5 class="accordion-button card-title pb-1" type="button" data-bs-toggle="collapse" data-bs-target="#flush-collapseOne" aria-expanded="true" aria-controls="flush-collapseOne">
                                Chọn nhân viên - Khu Vực:
                            </h5>
                            <hr class="mt-1"/>
                            <div id="flush-collapseOne" class="accordion-collapse collapse show" aria-labelledby="flush-headingOne" data-bs-parent="#accordionFlushExample" style="">
                                <ul id="sortable_dole" class="ps-0 d-flex flex-wrap">
                                    @foreach($areas as $area)
                                        <li data-value="{{$area->id}}" class="d-inline-block me-1 mb-1 pe-1 btn btn-{{ count($area->customers) > 0 ? 'primary' : 'secondary' }}">
                                            {{$area->name}} <span class="badge bg-{{ count($area->customers) > 0 ? 'white' : 'danger' }} text-{{ count($area->customers) > 0 ? 'primary' : 'white' }}">{{ count($area->customers) }}</span>
                                        </li>
                                    @endforeach
                                </ul>
                                <div class="row pb-2">
                                    <div class="col-md-3">
                                        <input type="number" name="from_row" onkeyup="if(parseInt(this.value) > 2000 || parseInt(this.value) < 1){ this.value = ''; return false; }"  class="from_row form-control" value="{{ old('from_row') }}" placeholder="Từ dòng số">
                                    </div>
                                    <div class="col-md-3 mt-sm-2 mt-md-0">
                                        <input type="number" name="to_row" value="{{ old('to_row') }}" onkeyup="if(parseInt(this.value) > 2000 || parseInt(this.value) < 1){ this.value = ''; return false; }" class="to_row form-control" placeholder="Đến số dòng">
                                    </div>
                                    <div class="col-md-2 ps-md-0 mt-sm-2 mt-md-0">
                                        <a href="#" id="btn-submit-customer" class="btn btn-success">Đăng Ký</a>
                                    </div>
                                </div>
                                @if ($errors->any())
                                    <div class="alert alert-danger alert-dismissible fade show mt-2" role="alert">
                                        @foreach ($errors->all() as $error)
                                            {{ $error }}
                                        @endforeach
                                        <button type="button" class="btn-close" data-bs-dismiss="alert" aria-label="Close"></button>
                                    </div>
                                @endif
                            </div>
                        </div>
                      </div>
                    </div>
                </div>
            </div>
        </div>

        <div class="row">
            <div class="card">
                <div class="card-body table-responsive">
                    <!-- Table with stripped rows -->
                    <table class="table table-striped">
                        <thead>
                            <tr>
                                <th scope="col">#</th>
                                <th scope="col">Chọn</th>
                                <th scope="col">Số Hợp đồng</th>
                                <th scope="col">Ngày Tham Gia</th>
                                <th scope="col">Mệnh Giá</th>
                                <th scope="col">Năm Đáo Hạn</th>
                                <th scope="col">Họ Và Tên</th>
                                <th scope="col">Giới Tính</th>
                                <th scope="col">Ngày Sinh</th>
                                <th scope="col">Điện Thoại</th>
                                <th scope="col">Tuổi</th>
                                <th scope="col text-center">Địa chỉ</th>
                                <th scope="col">Ngày Tạo</th>
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
                                <th class="customer_choose">
                                    <div class="form-check d-flex justify-content-center">
                                        <input name="choose_customers[]" class="form-check-input" type="checkbox" value="{{ $customer->id }}" id="flexCheckDefault">
                                        <label class="form-check-label" for="flexCheckDefault">
                                        </label>
                                    </div>
                                </th>
                                <td>{{ $customer->so_hop_dong }}</td>
                                <td>{{ @$customer->ngay_tham_gia }}</td>
                                <td>{{ is_numeric($customer->menh_gia) ? number_format($customer->menh_gia) : '' }}</td>
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
                            </tr>
                            @php
                            $j--;
                            @endphp
                            @endforeach
                        </tbody>
                    </table>
                    @if($customers->total() == 0)
                        <h5 class="text-center pt-5 pb-5"><b>ĐÃ CẤP HẾT DỮ LIỆU CHO CÁC KHU VỰC</b></h5>
                    @else
                        {!! $customers->links('_partials.pagination') !!} 
                    @endif
                </div>
            </div>
        </div>
   </form>
</section>

@endsection

@push('scripts')

<script>
$( function() {
    $( "#sortable_dole" ).sortable();
    $( "#sortable-dole" ).disableSelection();
} );

$("#sortable_dole li").click(function() {
    var itemChoose = '';
    if ($(this).hasClass('bg-success')) {
        $('#sortable_dole li:not(.bg-success)').removeClass('is-disabled');
        $(this).removeClass('bg-success');
    } else {
        $(this).addClass('bg-success');
        $('#sortable_dole li:not(.bg-success)').addClass('is-disabled');
    }
});

$( "#btn-submit-customer" ).click(function() {
    var from_row = $('.from_row').val() || 0;
    var to_row = $('.to_row').val() || 0;
   
    var checked = false;
    $( ".table .customer_choose input" ).each(function( index ) {
        if ($(this).is(':checked')) {
            checked = true;
        }
    })

    if (!checked) {
        if (from_row == '') {
            from_row = 0;
        }
        if (to_row == '') {
            to_row = 0;
        }
    }
    $( ".table .customer_choose input" ).each(function( index ) {
        if (index < Number(to_row) ) {
            $(this).prop('checked', true);
        } else {
            return;
        }
    });

    var value = $('#sortable_dole li').filter(function(){
        return $(this).hasClass("bg-success");
    });

    if ($( ".table .customer_choose input" ).length > 0) {
        if ($(value.get(0)).data('value')) {
            $('#form-customer').append(`<input type="hidden" name="area" value="${$(value.get(0)).data('value')}" />`);
        }
        $( "#form-customer" ).submit();
    } else {
        alert("Vui lòng import dữ liệu excel vào hệ thống!");
    }
});

</script>
@endpush
