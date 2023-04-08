<?php

namespace App\Http\Controllers;

use Illuminate\Http\Request;
use App\Http\Controllers\Controller;
use App\Models\Area;
use App\Models\User;
use App\Models\Customer;
use Brian2694\Toastr\Facades\Toastr;
use App\Models\AreaUser;
use App\Models\HistoryArea;
use DB;
use App\Http\Traits\Pagination;

class AreaController extends Controller
{
    public $_status = '';
    public $dataAreas = '';

    const APPOINTMENT = 0;
    const CUSTOMER_ACTIVE = 1;
    const USER_ACTIVE = 1;
    const AREA_ACTIVE = 1;
    const CALLED = 1;

    public function __construct()
    {
        $this->dataAreas = Area::select('name', 'id', 'note', 'created_at')->opening()->get();
    }

    public function index(Request $request)
    {
        $areas = Area::with('customers')->orderBy('updated_at', 'desc')->get();
        $areaAtatus = Area::getStatus();

        return view('area.list', compact('areas', 'areaAtatus'));
    }

    public function deleteAll() {
        DB::beginTransaction();
        try {
            foreach(Area::all() as $item) {
                $item->delete();
            }
            DB::commit();
            Toastr::success("Xoá khu vực thành công");
        } catch(\Exception $ex) {
            DB::rollback();
            Toastr::error("Xoá khu vực thất bại". $ex->getMessage());
        }
        return redirect()->route('index_area');
    }

    public function create(Request $request)
    {
        $validator = $request->validate([
            'name' => 'required|unique:areas|min:2',
        ],[
            'name.required' => 'Tên khu vực bắt buộc phải có',
            'name.unique' => 'Tên khu vực đã tồn tại',
            'name.min' => 'Tên khu vực quá ngắn'
        ]);

        $data = $request->all();
        $product = Area::create($data);
        if ($product) {
            Toastr::success("Tạo khu vực ". $request->get('name') ." thành công!");
        } else {
            Toastr::error("Tạo khu vực ". $request->get('name') ." thất bại!");
        }

        return redirect()->route('index_area');
    }

    public function edit($id, Request $request)
    {
        $area = Area::find($id);
        $areas = Area::all();
        $areaAtatus = Area::getStatus();

        return view('area.list', compact('area', 'areas', 'areaAtatus'));
    }

    public function postEdit($id, Request $request)
    {
        $model = Area::find($id);
        $model->name = $request->get('name');
        $model->status = $request->get('status');
        $model->note = $request->get('note');

        try {
            $model->update();
            Toastr::success("Cập nhật khu vực ". $model->name ." thành công!");
        } catch (\Exception $ex) {
            Toastr::error("Cập nhật khu vực ". $model->name ." thất bại!". $ex->getMessage());
        }

        return redirect()->route('index_area');
    }

    public function moveAreaBack(Request $request)
    {
        $area_id = $request->get('area_id');
        $user_id = $request->get('user_id');

        try {
            Area::where([
                'user_id' => $user_id,
                'id' => $area_id,
            ])->update([
                'user_id' => null
            ]);
        } catch (\Exception $ex) {
            return \Response::json(['message' => $ex->getMessage()], 400);
        }
        return \Response::json(['message' => 'Cập nhật thành công!'], 200);
    }

    public function permissionArea(Request $request)
    {
        $area_id = $request->get('area_id');
        $user_id = $request->get('user_id');
        DB::beginTransaction();
        try {
            if (Area::where('id', $area_id)->count() > 0) {
                Area::where('id', $area_id)->update([ 'user_id' => $user_id ]);
            }
            \DB::commit();
        } catch (\Exception $ex) {
            DB::rollback();
            return \Response::json(['message' => $ex->getMessage()], 400);
        }
        return \Response::json(['message' => 'Cập nhật thành công!'], 200);
    }

    public function delete($id, Request $request)
    {
        if (isset($_COOKIE['area_id']) && $id == $_COOKIE['area_id']) {
            unset($_COOKIE['area_id']);
            setcookie('area_id', null, -1, '/');
        }

        DB::beginTransaction();
        try {
            $area = Area::findOrFail($id);
            $area->delete();
            Toastr::success("Xoá khu vực ". $area->name ." thành công!");
            DB::commit();
        } catch (\Exception $ex) {
            DB::rollback();
            Toastr::error("Vui lòng mở cấp quyền! Rồi hãy xoá khu vực!". $ex->getMessage());
        }

        return redirect()->route('index_area');
    }

    public function doleCustomersToArea()
    {
        $areas = Area::with('customers')->orderBy('name', 'ASC')->get();
        $customers = Customer::whereNull('area_id')->whereNull('called')->paginate(2000);

        return view('area.list-dole', compact('areas', 'customers'));
    }

    public function postDoleCustomersToArea(Request $request)
    {
        $request->validate([
            'area' => 'required',
        ], [
            'area.required' => 'Vui lòng chọn khu vực.',
        ]);

        $ids = $request->input('choose_customers');

        if (empty($ids)) {
            Toastr::error("Khách hàng chưa được chọn! ");
            return redirect()->back();
        }
        $area = Area::findOrFail($request->input('area'));

        DB::beginTransaction();
        try {
            Customer::whereIn('id', $request->get('choose_customers'))
                    ->update(['area_id' => $request->input('area')]);
            DB::commit();
            Toastr::success("Đã thêm một số khách hàng vào khu vực ". $area->name);
        } catch (\Exception $ex) {
            DB::rollback();
            Toastr::error("Cấp quyền cho khu vực bị thất bại! ". $ex->getMessage());
        }
        return redirect()->back();
    }

    public function addAreaToUser(Request $request)
    {
        $areas = Area::whereNull('user_id')->where('status', self::AREA_ACTIVE)->orderBy('name', 'ASC')->get();
        $areaUsers = User::with('customers_area_has_users')->orderBy('username', 'ASC')->get();
        $numberCustomerArea = Area::with('customers')->whereNotNull('areas.user_id')->get();

        return view('area.add-area-to-user', [ 'areas' => $this->dataAreas, 'areaUsers' => $areaUsers, 'areas' => $areas, 'numberCustomerArea' => $numberCustomerArea ]);
    }

    public function reopenArea($id)
    {
        DB::beginTransaction();
        try {
             $numberRecord = Customer::where('area_id', $id)->where('called', self::CALLED)
                ->whereNotNull('type_call')
                ->where('updated_at', '>=', \Config::get('config.DATE_REOPEN'))
                ->where('type_call', '<>', self::APPOINTMENT)
                ->update(['type_call' => null, 'called' => null, 'comment' => null]);

            $area = Area::find($id);
            if ($numberRecord === 0) {
                Toastr::warning("Dữ liệu cần khôi phục đã hết trong khu vực ". $area->name);
            } else {
                $userID = $area->user_id;
                $area->user_id = null;
                $area->updated_at = \Carbon\Carbon::now();
                $area->save();

                // history
                HistoryArea::create([
                    'area_id' => $id,
                    'author_reopen' => \Auth::id(),
                    'count_record' => $numberRecord,
                    'user_id' => $userID
                ]);

                Toastr::success("Khôi phục thành công ". $numberRecord ." khách hàng trong khu vực ". $area->name);
            }
            DB::commit();
        } catch (\Exception $ex) {
            DB::rollback();
            Toastr::error("Khôi phục khu vực có lỗi xảy ra! Liên hệ SUPPORT ZALO(0964944719)". $ex->getMessage());
        }

        return redirect()->back();
    }
}
