<?php

namespace App\Http\Controllers;

use App\Models\Customer;
use Illuminate\Http\Request;
use App\Http\Controllers\Controller;
use Maatwebsite\Excel\Facades\Excel;
use Brian2694\Toastr\Facades\Toastr;
use App\Exports\AppointmentExport;
use App\Imports\CustomerImport;
use App\Models\HistoryExcel;
use Carbon\Carbon;

class ExcelController extends Controller
{
    const EXCEL_TYPE_FILE = '.xlsx';

    public $dataCustomers = '';

    public function import(Request $request)
    {
        $importHistory = HistoryExcel::with('user')->orderBy('created_at', 'desc')->paginate(20);


        return view('excel.list', compact('importHistory'));
    }

    public function deleteExcelCustomer($id, Request $request)
    {
        try {
            $customer = Customer::find($id);
            $customer->delete();
            Toastr::success("Xóa khách hàng ". $customer->ten ." thành công!");
        } catch (\Exception $ex) {
            Toastr::error("Xóa khách hàng ". $customer->ten ." thất bại!". $ex->getMessage());
        }

        return redirect()->route('data_import');
    }

    public function history(Request $request)
    {
        return view('excel.history');
    }

    public function postImport(Request $request)
    {
        $request->validate([
            'file' => 'required|max:2048|mimes:xlsx,xls',
        ],[
            'file.required' => 'Vui lòng chọn file Excel để import dữ liệu.',
            'file.max' => 'Tập tin quá lớn.',
            'file.mimes' => 'File không được cài mật khẩu, định dạng file không đúng. Chỉ cho phép import file Excel(xlsx,xls) thôi.'
        ]);

        $fileNameOrigin = $request->file('file')->getClientOriginalName();
        $import = new HistoryExcel();
        $import->user_id = \Auth::id();
        $import->file_name = $fileNameOrigin;
        $import->created_at = \Carbon\Carbon::now();

        \DB::beginTransaction();

        try {
            $numberRows = Customer::count();
            Excel::queueImport(new CustomerImport, request()->file('file'));
            $import->updated_at = \Carbon\Carbon::now();

            \DB::commit();
        } catch (\Maatwebsite\Excel\Validators\ValidationException $e) {
            $failures = $e->failures();
            $errormessage = "";

            foreach ($failures as $failure) {
                $errormess = "";
                foreach($failure->errors() as $error) {
                    $errormess = $errormess.$error;
                }
                $errormessage = $errormessage."\n Dòng số ".$failure->row().", ".$errormess."<br>";
            }
            $import->status = $errormessage;
            $import->updated_at = \Carbon\Carbon::now();
            $import->save();
            return \Response::json(['message' => $errormessage], 400);
        } catch (\Illuminate\Database\QueryException $e) {
            $import->status = $e->getMessage();
            $import->updated_at = \Carbon\Carbon::now();
            $import->save();
            $errorCode = $e->errorInfo[1];
            if($errorCode == 1062)
                \DB::rollback();
            return \Response::json(['message' => 'Dữ liệu thêm vào database đã có lỗi xảy ra.'. $e->getMessage()], 500);
        } catch(\Exception $ex) {
            return \Response::json(['message' => $ex->getMessage()], 500);
        }
        $numberRows1 = Customer::count();
        $countImported = $numberRows1 - $numberRows;
        if ($numberRows == $numberRows1 || $countImported < 0) {
            $import->status = "Trùng Lặp";
            $import->save();
            Toastr::warning('Vui lòng kiểm tra lại dữ liệu đã bị trùng lặp!');
        } else {
            $import->number = $countImported;
            $import->status = "Thành Công";
            $import->save();
            Toastr::success("Import thành công! ". $countImported ." dòng dữ liệu ");
        }

        $time = date('Y-m-d H:i:s');
        $time = str_replace(':', '_', $time);
        $time = str_replace(' ', '_', $time);

        $uploadedFile = $request->file('file');
        $filename = $time.$uploadedFile->getClientOriginalName();
        \Storage::disk('local')->putFileAs(
            'files/',
            $uploadedFile,
            $filename
        );

        return \Response::json(['message' => "Import thành công! ". $countImported ." dòng dữ liệu "], 200);
    }

    public function appointmentExport()
    {
        $time = date('Y-m-d H:i:s');
        $time = str_replace(':', '_', $time);
        $time = str_replace(' ', '_', $time);

        return Excel::download(new AppointmentExport, $time . 'danh-sach-da-hen.xlsx');
    }

    public function historyDelete($id, Request $request)
    {
        try {
            $historyExcel = HistoryExcel::findOrFail($id);
            $historyExcel->delete();
            Toastr::success("Xóa lịch sử import ". $historyExcel->file_name ." thành công!");
        } catch (\Exception $ex) {
            Toastr::error("Xóa lịch sử import thất bại!". $ex->getMessage());
        }

        return redirect()->route('data_import');
    }

    public function seachSHD(Request $request)
    {
        $checkExits = false;
        if (request('search')) {
            $checkExits = Customer::where('so_hop_dong', '=', $request->get('search'))->exists();
        } else {
            Toastr::info('Vui lòng nhập từ khóa tìm kiếm');
            return redirect()->back();
        }

        if($checkExits) {
            Toastr::error('Dữ liệu đã tồn tại trong hệ thống!');
        } else {
            Toastr::info('Dữ liệu chưa có trong hệ thống!');
        }

        return redirect()->back();
    }
}
