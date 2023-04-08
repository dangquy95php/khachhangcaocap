<?php

namespace Database\Seeders;

use Illuminate\Database\Seeder;

class CustomerSeeder extends Seeder
{
    /**
     * Run the database seeds.
     *
     * @return void
     */
    public function run()
    {
        
        $data = [
            // [
            //         'so_thu_tu' => '1',
            //         'vpbank' => 'bank1',
            //         'msdl' => 'msdl',
            //         'cv' => 'cv',
            //         'so_hop_dong' => '09223322',
            //         'menh_gia' => '1000000',
            //         'nam_dao_han' => '2040',
            //         'ho' => 'dang',
            //         'ten' => 'quy',
            //         'ten_kh' => 'khach hang 1',
            //         'gioi_tinh' => 'M',
            //         'ngay_sinh' => '1996',
            //         'tuoi' => '20',
            //         'dien_thoai' => '097292323',
            //         'dia_chi_cu_the' => 'dia_chi_cu_the1',
            //         'type_call' => 1,
            //         'called' => 1,
            //         'area_id' => 1,
            //         'comment' => 'comment1',
            //         'updated_at' => '2022-06-25 00:27:56',
            //         'date_max' => null
            //     ],
            //     [
            //         'so_thu_tu' => '1',
            //         'vpbank' => 'bank2',
            //         'msdl' => 'msdl',
            //         'cv' => 'cv',
            //         'so_hop_dong' => '0922332',
            //         'menh_gia' => '2000000',
            //         'nam_dao_han' => '2030',
            //         'ho' => 'dang',
            //         'ten' => 'quy2',
            //         'ten_kh' => 'khach hang 2',
            //         'gioi_tinh' => 'M',
            //         'ngay_sinh' => '1994',
            //         'tuoi' => '20',
            //         'dien_thoai' => '0973733',
            //         'dia_chi_cu_the' => 'dia_chi_cu_the2',
            //         'type_call' => 1,
            //         'area_id' => 1,
            //         'called' => 0,
            //         'comment' => 'comment2',
            //         'updated_at' => '2022-06-26 00:27:56',
            //         'date_max' => null
            //     ],
            //     [
            //         'so_thu_tu' => '1',
            //         'vpbank' => 'bank3',
            //         'msdl' => 'msdl',
            //         'cv' => 'cv',
            //         'so_hop_dong' => '09223321',
            //         'menh_gia' => '3000000',
            //         'nam_dao_han' => '2022',
            //         'ho' => 'dang',
            //         'ten' => 'quy1',
            //         'ten_kh' => 'khach hang 3',
            //         'gioi_tinh' => 'F',
            //         'ngay_sinh' => '1995',
            //         'tuoi' => '12',
            //         'dien_thoai' => '098723212',
            //         'dia_chi_cu_the' => 'dia_chi_cu_the3',
            //         'type_call' => 1,
            //         'area_id' => 2,
            //         'called' => 0,
            //         'comment' => 'comment3',
            //         'updated_at' => '2022-06-27 00:27:56',
            //         'date_max' => null
            //     ]
                ];
        \DB::table('customers')->insert($data);
    }
}
