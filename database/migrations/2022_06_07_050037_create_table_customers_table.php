<?php

use Illuminate\Database\Migrations\Migration;
use Illuminate\Database\Schema\Blueprint;
use Illuminate\Support\Facades\Schema;

class CreateTableCustomersTable extends Migration
{
    /**
     * Run the migrations.
     *
     * @return void
     */
    public function up()
    {
        Schema::dropIfExists('customers');
        Schema::create('customers', function (Blueprint $table) {
            $table->bigIncrements('id', true);
            $table->string('so_thu_tu')->nullable();
            $table->string('vpbank')->nullable();
            $table->string('msdl')->nullable();
            $table->string('cv')->nullable();
            $table->string('so_hop_dong')->unique();
            $table->string('menh_gia')->nullable();
            $table->string('nam_dao_han')->nullable();
            $table->string('ho')->nullable();
            $table->string('ten')->nullable();
            $table->string('ten_kh')->nullable();
            $table->string('gioi_tinh')->nullable();
            $table->string('ngay_sinh')->nullable();
            $table->string('tuoi')->nullable();
            $table->string('dien_thoai')->nullable();
            $table->string('dia_chi_cu_the')->nullable();
            $table->text('comment')->nullable();
            $table->integer('type_call')->nullable();
            $table->boolean('called')->nullable();
            $table->integer('area_id')->unsigned()->nullable();
            $table->timestamp('created_at')->useCurrent();
            $table->timestamp('updated_at')->useCurrent();
        });

        Schema::table('customers', function($table) {
            $table->foreign('area_id')->references('id')->on('areas');
        });
    }

    /**
     * Reverse the migrations.
     *
     * @return void
     */
    public function down()
    {
        Schema::dropIfExists('customers');
    }
}
