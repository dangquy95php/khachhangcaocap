<?php

namespace App\Imports;

use App\Models\Customer;
use Illuminate\Support\Collection;
use Maatwebsite\Excel\Concerns\ToCollection;
use Maatwebsite\Excel\Concerns\WithHeadingRow;
use Maatwebsite\Excel\Concerns\WithChunkReading;
use Illuminate\Contracts\Queue\ShouldQueue;
use Maatwebsite\Excel\Concerns\Importable;
use Maatwebsite\Excel\Concerns\WithStartRow;
use Maatwebsite\Excel\Concerns\SkipsEmptyRows;
use Maatwebsite\Excel\Concerns\WithCalculatedFormulas;
use Maatwebsite\Excel\Concerns\WithValidation;
use Maatwebsite\Excel\Concerns\WithEvents;
use Maatwebsite\Excel\Events\BeforeImport;

class CustomerImport implements ToCollection, SkipsEmptyRows, WithEvents, WithHeadingRow, WithCalculatedFormulas, WithChunkReading, WithValidation, ShouldQueue
{
    use Importable;

    const RUN_EVERY_TIME = 200;
    public $data = [];

    public function headingRow(): int
    {
        return 1;
    }

    public function collection(Collection $rows)
    {
        $allCustomers = Customer::select('so_hop_dong')->get();

        foreach ($rows as $row) 
        {
            $this->data[] = $row;

            foreach($allCustomers->chunk(200) as $item) {
                if (!empty($item->so_hop_dong)
                && !empty($row['so_hop_dong'])
                && $item->so_hop_dong == $row['so_hop_dong']) {
                    break;
                }
            }

            foreach(collect($this->data)->chunk(200) as $item) {
                if (!empty($item->so_hop_dong)
                && !empty($row['so_hop_dong'])
                && $item->so_hop_dong == $row['so_hop_dong']) {
                    break;
                }
            }

           try {
                Customer::create([
                    'so_thu_tu'        => @$row['so_thu_tu'],
                    'vpbank'           => @$row['vpbank'],
                    'msdl'             => @$row['msdl'],
                    'cv'               => @$row['cv'],
                    'so_hop_dong'      => @$row['so_hop_dong'],
                    'ngay_tham_gia'    => @$row['ngay_tham_gia'],
                    'menh_gia'         => @$row['menh_gia'],
                    'nam_dao_han'      => @$row['nam_dao_han'],
                    'ho'               => @$row['ho'],
                    'ten'              => @$row['ten'],
                    'ten_kh'           => @$row['ho'] .' '. @$row['ten'],
                    'gioi_tinh'        => @$row['gioi_tinh'],
                    'ngay_sinh'        => \PhpOffice\PhpSpreadsheet\Shared\Date::excelToDateTimeObject(@$row['ngay_sinh'])->format('d/m/Y'),
                    'tuoi'             => @$row['tuoi'],
                    'dien_thoai'       => @$row['dien_thoai'],
                    'dia_chi_cu_the'   => @$row['dia_chi_cu_the'],
                    'cccd'             => @$row['cccd'],
                ]);
                \Log::info($row);
            } catch (\Exception $ex) {

            }
        }
    }


    public function registerEvents(): array
    {
        return [
            BeforeImport::class => function (BeforeImport $event) {
                $totalRows = $event->getReader()->getTotalRows();
                $sheetName = key($totalRows);

                if (!empty($sheetName) && $totalRows[$sheetName] >= 10000) {
                   throw new \Exception('Số dòng trong file không được nhiều hơn 10000');
                }
            }
        ];
    }

    public function rules(): array
    {
        return [
            'so_hop_dong' => ['required'],//số hợp đồng
        ];
    }

    /**
     * @return array
     */
    public function customValidationMessages()
    {
        return [
            'so_hop_dong.required' => 'Không được để trống cột số :attribute. Xem lại dòng đầu tiên không được để trống.',
        ];
    }

    public function chunkSize(): int
    {
        return self::RUN_EVERY_TIME;
    }
}
