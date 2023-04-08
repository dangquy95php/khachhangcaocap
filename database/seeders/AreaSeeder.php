<?php

namespace Database\Seeders;

use Illuminate\Database\Seeder;

class AreaSeeder extends Seeder
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
            // [
            //     'name' => 'KV1',
            //     'status' => 1,
            //     'user_id' => 2,
            //     'note' => ''
            // ],
            // [
            //     'name' => 'KV2',
            //     'status' => 1,
            //     'user_id' => 1,
            //     'note' => ''
            // ],
            // [
            //     'name' => 'KV3',
            //     'status' => 1,
            //     'user_id' => 1,
            //     'note' => ''
            // ],

            // [
            //     'name' => 'MÙI',
            //     'status' => 1,
            //     'note' => ''
            // ],
            // [
            //     'name' => 'TÂM',
            //     'status' => 1,
            //     'note' => ''
            // ],
            // [
            //     'name' => 'KIM LUYẾN',
            //     'status' => 1,
            //     'note' => ''
            // ],
            // [
            //     'name' => 'HẰNG',
            //     'status' => 1,
            //     'note' => ''
            // ],
            // [
            //     'name' => 'VÂN ANH',
            //     'status' => 1,
            //     'note' => ''
            // ],
            // [
            //     'name' => 'VY',
            //     'status' => 1,
            //     'note' => ''
            // ],
            // [
            //     'name' => 'THOA',
            //     'status' => 1,
            //     'note' => ''
            // ],
            // [
            //     'name' => 'LUYÊN',
            //     'status' => 1,
            //     'note' => ''
            // ],
            // [
            //     'name' => 'HUỲNH LINH',
            //     'status' => 1,
            //     'note' => ''
            // ],
            // [
            //     'name' => 'MỸ',
            //     'status' => 1,
            //     'note' => ''
            // ],
            // [
            //     'name' => 'HUYỀN',
            //     'status' => 1,
            //     'note' => ''
            // ],
            // [
            //     'name' => 'LƯU TUYÊN',
            //     'status' => 1,
            //     'note' => ''
            // ],
            // [
            //     'name' => 'KV1',
            //     'status' => 1,
            //     'note' => '',
            // ],
            // [
            //     'name' => 'KV2',
            //     'status' => 1,
            //     'note' => '',
            // ]
        ];

        \DB::table('areas')->insert($data);
    }
}