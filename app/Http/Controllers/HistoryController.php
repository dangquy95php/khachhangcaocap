<?php

namespace App\Http\Controllers;

use Illuminate\Http\Request;
use App\Http\Controllers\Controller;
use App\Models\HistoryArea;

class HistoryController extends Controller
{
    
    public function indexArea(Request $request)
    {
        $data = HistoryArea::with('area', 'user', 'author')->orderBy('updated_at', 'desc')->get();

        return view('history.area', compact('data'));
    }
}
