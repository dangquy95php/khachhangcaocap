<?php

namespace App\Http\Controllers;

use Auth;
use Illuminate\Http\Request;
use App\Http\Controllers\Controller;
use App\Http\Requests\User\LoginUserRequest;
use App\Http\Requests\User\CreateUserRequest;
use App\Http\Requests\User\UpdateUserRequest;
use App\Models\Area;
use App\Models\User;
use Brian2694\Toastr\Facades\Toastr;
use Hash;
use Artisan;

class UserController extends Controller
{
    const USER_ROLE = 1;
    const ADMIN_ROLE = 2;
    const USER_ACTIVED = 1;

    public function index()
    {
        if (Auth::check()) {
            return redirect()->route('home');
        }
        return redirect()->route('login');
    }

    public function login(Request $request)
    {
        if (Auth::check()) {
            return redirect()->route('home');
        }
        return view('account.login');
    }

    public function postLogin(LoginUserRequest $request)
    {
        $data = [
            'username' => $request->get('username'),
            'password' => trim($request->get('password')),
            'status' => self::USER_ACTIVED,
        ];
        \Log::info($data);

        $remember_me = $request->has('remember') ? true : false;

        if (Auth::attempt($data, $remember_me)) {
            Toastr::success('Đăng nhập vào hệ thống thành công!');

            if (Auth::user()->role == self::USER_ROLE) {
                return redirect()->route('home');
            } else {
                return redirect()->route('dashboard');
            }
        } else {
            Toastr::error('Tên người dùng hoặc mật khẩu không chính xác!');
            return redirect()->route('login');
        }
        return redirect()->route('home');
    }

    public function logout(Request $request)
    {
        Auth::logout();
        Toastr::success('Đăng xuất thành công!');
        return redirect()->route('login');
    }

    public function create(Request $request)
    {
        return view('account.create');
    }

    public function postCreate(CreateUserRequest $request)
    {
        $data = [
            'name'          => $request->input('name'),
            'username'      => $request->input('username'),
            'email'         => $request->input('email'),
            'status'        => $request->input('status'),
            'role'          => $request->input('role'),
            'password'      => $request->input('password'),
        ];
        try {
            User::create($data);
            Toastr::success('Tạo tài khoản '. $request->input('username') .' thành công');
        } catch(\Exception $e) {
            Toastr::success('Tạo tài khoản thất bại!'. $e->getMessage());
        }

        return redirect()->route('list_account');
    }

    public function list(Request $request)
    {
        $data = User::select('id', 'name', 'email', 'username', 'role', 'status', 'created_at')
                        ->orderBy('created_at', 'desc')->get();

        return view('account.list', compact('data'));
    }

    public function edit($id, Request $request)
    {
        $data = User::find($id);

        return view('account.create', compact('data'));
    }

    public function postEdit($id, UpdateUserRequest $request)
    {
        $user = User::find($id);
        $user->name = $request->get('name');
        $user->email = $request->get('email');
        $user->role = $request->get('role');
        $user->status = $request->input('status');

        if (!Hash::check($request->input('password'), $user->password) && $request->get('check_password') !== 'on') {
            $user->password = trim($request->input('password'));
            User::find($id)->update(['password' => $user->password]);
        }
        try {
            
            if($user->isDirty()) {
                Toastr::success('Thông tin người dùng đã thay đổi thành công.');
            } else {
                Toastr::warning('Dữ liệu chưa được thay đổi');
            }
            $user->save();
        } catch (\Exception $ex) {
            Toastr::error('Có lỗi khi lưu dữ liệu '. $ex->getMessage());
        }

        return redirect()->route('list_account');
    }

    public function delete($id, Request $request)
    {
        try {
            $user = User::find($id);
            Area::where(['user_id' => $user->id])->update([ 'user_id' => null ]);
            $user->delete();
            Toastr::success("Xóa nhân viên ". $user->username ." thành công!");
        } catch (\Exception $ex) {
            Toastr::error("Vui lòng bỏ cấp quyền. Xóa nhân viên ". $user->username ." thất bại!". $ex->getMessage());
        }

        return redirect()->route('list_account');
    }
}
