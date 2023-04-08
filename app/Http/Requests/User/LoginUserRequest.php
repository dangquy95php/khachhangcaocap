<?php

namespace App\Http\Requests\User;

use Illuminate\Foundation\Http\FormRequest;

class LoginUserRequest extends FormRequest
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
            'username'    => 'required|max:80|min:2',
            'password'    => 'required|max:64|min:3',
        ];
    }

    public function messages()
    {
        return [
            'username.max'              => 'Tên người dùng nhập quá dài',
            'username.min'              => 'Tên người dùng quá ngắn',
            'username.required'         => 'Tên người dùng bắt buộc phải nhập',
            'password.max'              => 'Mật khẩu quá dài',
            'password.min'              => 'Mật khẩu quá ngắn',
            'password.required'         => 'Mật khẩu bắt buộc phải nhập',
        ];
    }
}
