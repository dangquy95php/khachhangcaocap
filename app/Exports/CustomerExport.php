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

class CustomerExport implements FromCollection, WithHeadings, ShouldAutoSize, WithEvents
{
    public function __construct(array $data)
    {
        $this->data = $data;
    }

    public function collection()
    {
        $start_date = $this->data['start_date'];
        $end_date = $this->data['end_date'];

        return Customer::get_data_export($start_date, $end_date);
    }

    public function headings() :array {

    	return ["Số thứ tự", "Vpbank", "Msdl", "CV", "Số hợp đồng", "Mệnh giá", "Năm đáo hạn", "Họ", "Tên", "Tên Khách Hàng", "Giới tính", "Ngày Sinh", "Tuổi", "Điện thoại", "Địa chỉ", "Ngày tham gia", "CCCD"];
    }

    public function registerEvents(): array
    {
        return [
            AfterSheet::class    => function(AfterSheet $event) {
                $cellRange = 'A1:Q1'; // All headers
                $event->sheet->getDelegate()->getStyle($cellRange)->getFont()->setSize(13);
                $event->sheet->getDelegate()->getStyle($cellRange)->getFill()
                ->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID)
                ->getStartColor()->setARGB('E3591B');
            },
        ];
    }
}
