<?php
namespace App\Http\Traits;
use Cache;

trait Pagination {

    public static function clearCache($idUser) {

        for($i = 1; $i <= 100; $i++) {
            if (Cache::has($idUser . $i)) {
                Cache::forget($idUser . $i);
            } else {
                break;
            }
        }
    }
}
