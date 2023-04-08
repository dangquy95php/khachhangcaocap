@section('title','Trang chủ')
@extends('layouts.template')

@section('breadcrumb')
  
   <h1>DANH SÁCH EXCEL IMPORT NGƯỜI DÙNG</h1>

    {{ Breadcrumbs::render('data_import') }}
    
@endsection

@section('content')

<section class="section">

    <div class="row">
        <div class="col-lg-12">
            <div class="card">
                <div class="card-body pt-3">
                    <!-- General Form Elements -->
                    <form>
                        <div class="row mb-3">
                            <div class="col-12">
                                <label for="inputText" class="col-sm-2 col-form-label">Chọn file Excel Import</label>
                                <input class="form-control" type="file" id="formFile">
                            </div>
                        </div>
                        <div class="row mb-3">
                            <div class="col-12">
                                <button type="submit" class="btn btn-primary">Improt File</button>
                                <button type="submit" class="btn btn-success text-white">Export File</button>
                            </div>
                        </div>
                    </form>
                    <!-- End General Form Elements -->
                </div>
            </div>
        </div>
    </div>

   <div class="row">
      <div class="col-lg-12">
         <div class="card">
            <div class="card-body">
               <!-- Table with stripped rows -->
               <table class="table table-striped">
                  <thead>
                     <tr>
                        <th scope="col">#</th>
                        <th scope="col">Số hợp đồng</th>
                        <th scope="col">Ngày tham gia</th>
                        <th scope="col">Số tiền</th>
                        <th scope="col">Ngày Đáo Hạn</th>
                        <th scope="col">Họ Và Tên</th>
                        <th scope="col">Giới Tính</th>
                        <th scope="col">Ngày Sinh</th>
                        <th scope="col">Địa chỉ</th>
                     </tr>
                  </thead>
                  <tbody>
                     <tr>
                        <th scope="row">1</th>
                        <td>73956316</td>
                        <td>20170822</td>
                        <td>50,000,000 </td>
                        <td>2027/08/18</td>
                        <td>NGUYỄN THẾ TÔN</td>
                        <td>Nam</td>
                        <td>12/16/79</td>
                        <td>C2.502 C/CƯ EHOME 4 BẮC SÀ ĐƯỜNG VĨNH PHÚ 41 P. VĨNH PHÚ, TX. THUẬN AN TỈNH BÌNH DƯƠNG</td>
                     </tr>
                     <tr>
                        <th scope="row">1</th>
                        <td>73956316</td>
                        <td>20170822</td>
                        <td>50,000,000 </td>
                        <td>2027/08/18</td>
                        <td>NGUYỄN THẾ TÔN</td>
                        <td>Nam</td>
                        <td>12/16/79</td>
                        <td>C2.502 C/CƯ EHOME 4 BẮC SÀ ĐƯỜNG VĨNH PHÚ 41 P. VĨNH PHÚ, TX. THUẬN AN TỈNH BÌNH DƯƠNG</td>
                     </tr>
                  </tbody>
               </table>
               <!-- End Table with stripped rows -->
            </div>
         </div>
      </div>
   </div>
</section>
@endsection