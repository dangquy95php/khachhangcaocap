<?php

namespace App\Exports;

use App\Models\Customer;
use Maatwebsite\Excel\Concerns\FromCollection;
use Maatwebsite\Excel\Concerns\WithHeadings;
use Maatwebsite\Excel\Concerns\ShouldAutoSize;
use Maatwebsite\Excel\Concerns\WithEvents;
use Maatwebsite\Excel\Events\AfterSheet;
use Carbon\Carbon;
use App\Models\User;

class AppointmentExport implements FromCollection, WithHeadings, ShouldAutoSize, WithEvents
{
    public function collection()
    {
        $result = [];
        $todayData = User::with(['get_data_today' => function($query) {
            return $query->select('so_hop_dong', 'ten_kh', 'dien_thoai', 'comment', 'gioi_tinh', 'tuoi', 'dia_chi_cu_the', 'customers.updated_at');
        }])->get();

        foreach($todayData as $data) {
           foreach($data->get_data_today as &$item) {
              $item->username = $data->name ?: $data->username;
              $item->gioi_tinh = ($item->gioi_tinh == 'M' ? 'Nam' : 'Nữ');
              $result[] = $item;
           }
        }

        return collect($result)->sortByDesc('updated_at');
    }

    public function headings() :array {
    	return ["Số hợp đồng", "Tên KH", "Số Điện Thoại", "Ghi Chú Khách Hàng", "Giới Tính", "Tuổi", "Địa chỉ cụ thể", "Thời gian gọi", "ID Nhân Viên", "Tài khoản gọi"];
    }

    public function registerEvents(): array
    {
        return [
            AfterSheet::class    => function(AfterSheet $event) {
                $cellRange = 'A1:J1'; // All headers
                $event->sheet->getDelegate()->getStyle($cellRange)->getFont()->setSize(13);
                $event->sheet->getDelegate()->getStyle($cellRange)->getFill()
                ->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID)
                ->getStartColor()->setARGB('E3591B');
            },
        ];
    }
}
