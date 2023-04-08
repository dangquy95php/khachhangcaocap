<?php

namespace App\Http\Requests\User;

use Illuminate\Foundation\Http\FormRequest;

class CreateUserRequest extends FormRequest
{
    /**
     * Determine if the user is authorized to make this request.
     *
     * @return bool
     */
    public function authorize()
    {
        return true;
    }

    /**
     * Get the validation rules that apply to the request.
     *
     * @return array
     */
    public function rules()
    {
        return [
            'password'                 => 'required|max:64|min:3|without_spaces',
            'password_confirmation'    => 'required|min:3|max:64|required_with:password|same:password',
            'username'                 => 'required|max:64|min:2|without_spaces|unique:users,username',
        ];

        // if (!$this->route('id')) {
        //     $validation = [
        //         'username'     => 'required|max:64|min:2|without_spaces|unique:users,username',
        //     ];
        // }

        // return $validation;
    }

    public function messages()
    {
        return [
            'password_confirmation.required' => 'Mật khẩu nhập lại bắt buộc phải nhập',
            'password_confirmation.same'     => 'Nhập lại mật khẩu không khớp với mật khẩu trên',
            'username.without_spaces'        => 'Tên người dùng không được nhập khoảng trắng',
            'password.without_spaces'        => 'Mật khẩu không được nhập khoảng trắng',
            'username.max'                   => 'Tên người dùng nhập quá dài',
            'username.min'                   => 'Tên người dùng quá ngắn',
            'username.required'              => 'Tên người dùng bắt buộc phải nhập',
            'password.max'                   => 'Mật khẩu quá dài',
            'password.min'                   => 'Mật khẩu quá ngắn',
            'password.required'              => 'Mật khẩu bắt buộc phải nhập',
            'username.unique'                => 'Tên người dùng đã tồn tại trong hệ thống'
        ];

        // if (!$this->route('id')) {
        //     $validation_mesage = [
        //         'username.unique'                => 'Tên người dùng đã tồn tại trong hệ thống'
        //     ];
        // }

        // return $validation_mesage;
    }
}
