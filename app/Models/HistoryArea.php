<?php

namespace App\Models;

use Illuminate\Database\Eloquent\Factories\HasFactory;
use Illuminate\Database\Eloquent\Model;

class HistoryArea extends Model
{
    use HasFactory;

    protected $table = 'history_area';

    /**
     * The attributes that are mass assignable.
     *
     * @var array
     */
    protected $fillable = [
        'area_id',
        'user_id',
        'count_record',
        'author_reopen',
    ];

    public function area()
    {
        return $this->hasOne(Area::class, 'id', 'area_id');
    }

    public function author()
    {
        return $this->hasOne(User::class, 'id', 'author_reopen');
    }

    public function user()
    {
        return $this->hasOne(User::class, 'id', 'user_id');
    }
}
