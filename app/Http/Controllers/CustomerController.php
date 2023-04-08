<?php

namespace App\Http\Controllers;

use Illuminate\Http\Request;
use App\Http\Controllers\Controller;
use App\Exports\CustomerExport;
use App\Imports\CustomerImport;
use Maatwebsite\Excel\Facades\Excel;
use Brian2694\Toastr\Facades\Toastr;
use App\Models\Customer;
use Cache;

class CustomerController extends Controller
{
    public $dataCustomers = '';
    const CACHE_EXPIRED = 600;

    public function __construct()
    {
        $this->dataCustomers = Customer::orderBy("updated_at", "desc")->paginate(20);
    }

    public function index(Request $request)
    {
        return view('customer.list', ['customers' => $this->dataCustomers]);
    }

    public function search(Request $request)
    {
        return view('customer.search');
    }

    public function export()
    {
        return Excel::download(new CustomerExport, 'customer.xlsx'); //download file export
        // return Excel::store(new CustomerExport, 'customer.xlsx', 'disk-name'); //lưu file export trên ổ cứng
    }

    public function delete(Request $request)
    {
        $request->validate([
            'start_date' => 'required',
            'end_date' => 'required',
        ], [
            'start_date.required' => 'Vui lòng chọn này bắt đầu!',
            'end_date.required' => 'Vui lòng chọn ngày kết thúc!'
        ]);

        switch ($request->input('action')) {
            case 'delete':
                $startDate = $request->get('start_date');
                $endDate = \Carbon\Carbon::parse($request->get('end_date'))->addDays(1);

                try {
                    $count = Customer::where('created_at', '>=', $startDate)->where('created_at', '<=', $endDate)->count();
                    Customer::where('created_at', '>=', $startDate)->where('created_at', '<=', $endDate)->delete();
                    if ($count == 0) {
                        Toastr::warning("Dữ liệu không được tìm thấy. Không thể xoá được!");
                    } else {
                        Toastr::success("Xoá tổng cộng ". $count ." thành công");
                    }
                } catch (\Exception $ex) {
                    Toastr::error("Xoá dữ liệu thất bại". $ex->getMessage());
                }
                return redirect()->back();
            break;

            case 'export':
                $time = date('Y-m-d H:i:s');
                $time = str_replace(':', '_', $time);
                $time = str_replace(' ', '_', $time);

                return Excel::download(new CustomerExport($request->all()), $time .'-danh-sach-export.xlsx'); //download file export
                break;
        }
    }

    public function deleteById($id, Request $request)
    {
        Customer::findOrFail($id)->delete();
        Toastr::success("Xoá dữ liệu thành công");

        return redirect()->back();
    }
}
