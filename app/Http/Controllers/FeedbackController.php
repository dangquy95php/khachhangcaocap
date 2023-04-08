<?php

namespace App\Http\Controllers;

use Illuminate\Http\Request;
use App\Http\Controllers\Controller;
use App\Models\Feedback;
use Brian2694\Toastr\Facades\Toastr;
use App\Http\Requests\User\FeedbackRequest;

class FeedbackController extends Controller
{
    const QUESTION = [
        'Phần mềm quản lý nhân viên sales giao diện trải nghiệm có tốt hơn phần mềm cũ của bạn không?',
        'Tốc độ load dữ liệu có nhanh không?',
        'Có khó khăn trong việc sử dụng phần mềm không?',
        'Bạn đánh giá trải nghiệm với phần mềm của chúng tôi như thế nào?',
        'Bạn có muốn nâng cấp thêm nhiều chức năng cho phần mềm này trong tương lai không?',
        'Bạn yêu thích điều gì nhất về phần mềm nhân viên sales chúng tôi?',
        'Bạn có cảm thấy thoải mái khi mua phần mềm của chúng tôi?',
        'Làm thế nào để chúng tôi có thể phục vụ bạn tốt hơn?',
    ];

    public function feedback(Request $request)
    {
        return view('feedback.index');
    }
    public function postFeedback(FeedbackRequest $request)
    {
        $k = "";
        foreach(self::QUESTION as $key => $item) {
            (string)$k = $key + 1;

            Feedback::create([
                'user_id' => \Auth::id(),
                'question_id' => $key+1,
                'question_content' => $item,
                'answer_content' => $request->get('question'. $k),
            ]);
        }

        \Cookie::queue('feedback', true, 60*20);
        return redirect()->route('home');
    }

    public function listFeedback(Request $request) {
        $data = Feedback::with('user')->get();

        return view('feedback.list', compact('data'));
    }
}
