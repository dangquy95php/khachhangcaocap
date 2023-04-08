<?php

namespace App\Models;

use Illuminate\Database\Eloquent\Model;

class Customer extends Model
{
    protected static $orderByColumn = 'created_at';

    protected static $orderByColumnDirection = 'desc';

    const BY_AREA_ACTIVE = 1; // đã update sau khi import. Tức là đã cho khách hàng vào khu vực

    const NOT_CALL = null;
    const NEW_CUSTOMER = 0;
    const AREA_ACTIVE = 1;
    /**
     * The attributes that are mass assignable.
     *
     * @var array
     */
    protected $fillable = [
        'id',
        'so_thu_tu',
        'vpbank',
        'ngay_tham_gia',
        'msdl',
        'cv',
        'so_hop_dong',
        'menh_gia',
        'nam_dao_han',
        'ho',
        'ten',
        'ten_kh',
        'gioi_tinh',
        'ngay_sinh',
        'tuoi',
        'dien_thoai',
        'dia_chi_cu_the',
        'comment',
        'type_call',
        'called',
        'date_max',
        'cccd',
    ];

    const APPOINTMENT = 0;
    const AGENT_STILL_CARE = 1;
    const LESS_MONEY = 2;
    const CALL_LATER = 3;
    const WRONG_OR_BUSY = 4;
    const CANCEL_DURATION = 5;
    const NEW_CONTRACT = 6;
    const OTHER = 7;

    const INFOR_OPTION = [
        self::APPOINTMENT => 'Đã hẹn',
        self::AGENT_STILL_CARE => 'Đại lý vẫn chăm sóc',
        self::LESS_MONEY => 'Khách hàng ít tiền',
        self::CALL_LATER => 'Khách hàng suy nghĩ, gọi lại sau',
        self::WRONG_OR_BUSY => 'KNM / Bận / Tắt máy / Sai số / Đổi số',
        self::CANCEL_DURATION => 'Hợp đồng Hủy / Đáo hạn / Hoàn trả giá trị',
        self::NEW_CONTRACT => 'Đã tham gia hợp đồng mới / Không tham gia',
        self::OTHER => 'Khác',
    ];

    public static function getInforOption()
    {
        return self::INFOR_OPTION;
    }

    public function getFullNameAttribute()
    {
        return $this->first_name . ' ' . $this->last_name;
    }

    public function area()
    {
        return $this->hasOne(Area::class, 'id');
    }

    public static function get_data_export($start, $end)
    {
        $start = $start . ' 00:00:00';
        $end = $end . ' 23:59:59';
        
        return Customer::where('created_at', '>=', $start)
                ->where('created_at', '<=', $end)
                ->where('called', self::NOT_CALL)
                ->orderBy('created_at', 'DESC')
                ->select('so_thu_tu', 'vpbank', 'msdl', 'cv', 'so_hop_dong', 'menh_gia', 'nam_dao_han', 'ho', 'ten', 'ten_kh', 'gioi_tinh', 'ngay_sinh', 'tuoi', 'dien_thoai', 'dia_chi_cu_the', 'ngay_tham_gia', 'cccd')->get();
    }
}
