<?php

namespace App\Http\Controllers;

use Illuminate\Http\Request;
use App\Http\Controllers\Controller;
use App\Models\Customer;
use Carbon\Carbon;
use App\Models\User;

class DashboardController extends Controller
{
   const APPOINTMENT = 0;
   const CALLED = 1;
   const DATE_MAX_TOTAL_CALLED = 1;

   public function dashboard(Request $request)
   {
      $result = [];
      $todayData = User::with('get_data_today')->get();
      foreach($todayData as $data) {
         foreach($data->get_data_today as &$item) {
            $item->username = $data->name ?: $data->username;
            $result[] = $item;
         }
      }
      $result = collect($result)->sortByDesc('updated_at');
      $totalCallCustomer = Customer::where('updated_at', '>=', \Carbon\Carbon::today())->where('called', '=', self::CALLED)->count();

      return view('dashboard', compact('result','totalCallCustomer'));
   }
}
