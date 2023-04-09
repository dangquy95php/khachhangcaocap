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

    const RUN_EVERY_TIME = 1000;
    public $data = [];

    public function headingRow(): int
    {
        return 1;
    }

    public function collection(Collection $rows)
    {
        foreach ($rows as $row) {
            if (!empty(@$row['so_hop_dong'])) {
                $isExists = Customer::where('so_hop_dong', @$row['so_hop_dong'])->exists();
                
                if (!$isExists) {
                    $result = Customer::where('ho', @$row['ho'])
                                    ->where('ten', @$row['ten'])
                                    ->where('gioi_tinh', @$row['gioi_tinh'])
                                    ->where('dien_thoai', @$row['dien_thoai'])
                                    ->where('tuoi', @$row['tuoi'])->exists();
                                  
                    if (!$result) {
                        try {
                            $tenKh = @$row['ten_kh'] ? @$row['ten_kh'] : @$row['ho'] .' '. @$row['ten'];
                            $customer = new Customer;
                            $customer->so_thu_tu = @$row['so_thu_tu'];
                            $customer->vpbank = @$row['vpbank'];
                            $customer->msdl = @$row['msdl'];
                            $customer->cv = @$row['cv'];
                            $customer->so_hop_dong = @$row['so_hop_dong'];
                            $customer->ngay_tham_gia = @$row['ngay_tham_gia'];
                            $customer->menh_gia = @$row['menh_gia'];
                            $customer->nam_dao_han = @$row['nam_dao_han'];
                            $customer->ho = @$row['ho'];
                            $customer->ten = @$row['ten'];
                            $customer->ten_kh = $tenKh;
                            $customer->gioi_tinh = @$row['gioi_tinh'];
                            $customer->ngay_sinh = \PhpOffice\PhpSpreadsheet\Shared\Date::excelToDateTimeObject(@$row['ngay_sinh'])->format('d/m/Y');
                            $customer->tuoi = @$row['tuoi'];
                            $customer->dien_thoai = @$row['dien_thoai'];
                            $customer->dia_chi_cu_the = @$row['dia_chi_cu_the'];
                            $customer->cccd = @$row['cccd'];
    
                            $customer->save();
                        } catch (\Exception $ex) { }
                    }
                }
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
