<?php

namespace App\Http\Controllers;

use App\Http\Controllers\Controller;
use Brian2694\Toastr\Facades\Toastr;
use Illuminate\Http\Request;
use App\Models\User;

class ReportController extends Controller
{
    public function index(Request $request)
    {
        $listCallOfStaff = User::with('customers_today_called')->orderBy('username', 'asc')->get();

        return view('report.index', compact('listCallOfStaff'));
    }

    public function ratings(Request $request)
    {
        $time = \Carbon\Carbon::now()->format('H:i:s');

        if ($time > '17:19:50' && $time < '17:20:10') {
            try {
                $ratings = User::with('customers_today_called')->orderBy('username', 'asc')->get();
                $result = [];
                foreach($ratings as &$datas) {
                    $appointment = 0;
                    $object = new \stdClass();

                    foreach($datas->customers_today_called as $item) {
                        if ($item->type_call == 0) {
                            $appointment++;
                        }
                    }
                    $object->name = $datas->name;
                    $object->appointment = $appointment;
                    $object->appointment_not_yet = count($datas->customers_today_called) - $appointment;

                    if ($object->appointment != 0 || $object->appointment_not_yet != 0) {
                        $result[] = $object;
                    }
                }
            } catch (\Exception $ex) {
                return \Response::json(['data' => 'Có lỗi xảy ra'. $ex->getMessage()], 500);
            }

            $result = collect($result)->sortByDesc( function($item) {
                if ($item->appointment != 0) {
                    return [
                        $item->appointment, - $item->appointment_not_yet
                    ];
                } else {
                    return [
                        -$item->appointment, $item->appointment_not_yet
                    ];
                }
            });

            return \Response::json(['data' => $result->values()->all()], 200);
        }

        return \Response::json(['data' => 'Lỗi thời gian'], 500);
    }
}
