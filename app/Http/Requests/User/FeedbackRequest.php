<?php

namespace App\Http\Requests\User;

use Illuminate\Foundation\Http\FormRequest;

class FeedbackRequest extends FormRequest
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
            'question1'    => 'required',
            'question2'    => 'required',
            'question3'    => 'required',
            'question4'    => 'required',
            'question5'    => 'required',
            'question6'    => 'required',
            'question7'    => 'required',
            'question8'    => 'required',
        ];
    }

    public function messages()
    {
        return [
           'question1.required' => 'Câu hỏi số 1 bất buộc phải nhập',
           'question2.required' => 'Câu hỏi số 2 bất buộc phải nhập',
           'question3.required' => 'Câu hỏi số 3 bất buộc phải nhập',
           'question4.required' => 'Câu hỏi số 4 bất buộc phải nhập',
           'question5.required' => 'Câu hỏi số 5 bất buộc phải nhập',
           'question6.required' => 'Câu hỏi số 6 bất buộc phải nhập',
           'question7.required' => 'Câu hỏi số 7 bất buộc phải nhập',
           'question8.required' => 'Câu hỏi số 8 bất buộc phải nhập',
        ];
    }
}
