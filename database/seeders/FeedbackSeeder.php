<?php

namespace Database\Seeders;

use Illuminate\Database\Seeder;

class FeedbackSeeder extends Seeder
{
    /**
     * Run the database seeds.
     *
     * @return void
     */
    public function run()
    {
        //
        $data = [
        //    [
        //     'user_id' => null,
        //     'question_content' => 'Phần mềm quản lý nhân viên sales giao diện trải nghiệm có tốt hơn phần mềm cũ của bạn không?',
        //     'question_id' => 1,
        //     'answer_content' => '',
        //    ],[
        //     'user_id' => null,
        //     'question_content' => 'Tốc độ load dữ liệu có nhanh không?',
        //     'question_id' => 2,
        //     'answer_content' => '',
        //    ],
        //    [
        //     'user_id' => null,
        //     'question_content' => 'Có khó khăn trong việc sử dụng phần mềm không?',
        //     'question_id' => 3,
        //     'answer_content' => '',
        //    ],
        //    [
        //     'user_id' => null,
        //     'question_content' => 'Bạn đánh giá trải nghiệm với phần mềm của chúng tôi như thế nào?',
        //     'question_id' => 4,
        //     'answer_content' => '',
        //    ],
        //    [
        //     'user_id' => null,
        //     'question_content' => 'Bạn có muốn nâng cấp thêm nhiều chức năng cho phần mềm này trong tương lai không?',
        //     'question_id' => 5,
        //     'answer_content' => '',
        //    ],
        //    [
        //     'user_id' => null,
        //     'question_content' => 'Bạn yêu thích điều gì nhất về phần mềm nhân viên sales chúng tôi?',
        //     'question_id' => 6,
        //     'answer_content' => '',
        //    ],
        //    [
        //     'user_id' => null,
        //     'question_content' => 'Bạn có cảm thấy thoải mái khi mua phần mềm của chúng tôi?',
        //     'question_id' => 7,
        //     'answer_content' => '',
        //    ],
        //    [
        //     'user_id' => null,
        //     'question_content' => 'Làm thế nào để chúng tôi có thể phục vụ bạn tốt hơn?',
        //     'question_id' => 8,
        //     'answer_content' => '',
        //    ],
        ];

        \DB::table('feedbacks')->insert($data);
    }
}