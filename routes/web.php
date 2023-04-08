<?php

use Illuminate\Support\Facades\Route;

/*
|--------------------------------------------------------------------------
| Web Routes
|--------------------------------------------------------------------------
|
| Here is where you can register web routes for your application. These
| routes are loaded by the RouteServiceProvider within a group which
| contains the "web" middleware group. Now create something great!
|
*/

Route::get('/', 'UserController@index');
Route::get('/login', 'UserController@login')->name('login');
Route::post('/login', 'UserController@postLogin')->name('post_login');

Route::group(['middleware' => 'auth'], function () {
    Route::get('/call/{area_id?}', 'HomeController@index')->name('home');
    Route::post('/call', 'HomeController@updateCusomter');
    Route::get('/logout', 'UserController@logout')->name('logout');
    Route::get('/feedback-list', 'FeedbackController@listFeedback');
});

Route::group(['middleware' => 'admin', 'prefix' => 'admin'], function () {

});

Route::group(['middleware' => 'auth', 'prefix' => 'customer'], function () {
    // Route::get('/call/{area_id?}', 'HomeController@index')->name('home');
    // Route::post('/call', 'HomeController@updateCusomter');

    Route::get('/', 'CustomerController@index')->name('list_customer');
    Route::get('/export-data', 'CustomerController@exportData')->name('customer.export-data');
    Route::get('/{id}/edit', 'HomeController@editCustomer')->name('customer.edit');
    Route::post('{id}', 'HomeController@postEditCustomer')->name('customer.edit.post');

    Route::get('/{id}/delete', 'CustomerController@deleteById')->name('customer.delete.byId');
    Route::get('/search', 'CustomerController@search')->name('search_customer');
    Route::get('/delete', 'CustomerController@delete')->name('customer.delete');
    Route::get('/{id}/edit', 'HomeController@editCustomer')->name('customer.edit');
    Route::post('/{id}', 'HomeController@postEditCustomer')->name('customer.edit.post');


    Route::get('/export/', 'CustomerController@export')->name('export_customer');
    Route::post('/import/', 'CustomerController@import')->name('import_customer');

    Route::get('/detail/{id}', 'HomeController@detail')->name('customer_detail');
    Route::post('/update', 'HomeController@update')->name('customer_update');
    Route::post('/save', 'HomeController@save')->name('customer_save');
});

Route::group(['middleware' => 'admin', 'prefix' => 'admin'], function () {
    Route::get('/create', 'UserController@create')->name('create_account');
    Route::post('/create', 'UserController@postCreate')->name('post_create_account');

    Route::get('/', 'DashboardController@dashboard')->name('dashboard');

    Route::group(['prefix' => 'area'], function () {
        Route::get('/', 'AreaController@index')->name('index_area');
        Route::get('/delete/all', 'AreaController@deleteAll')->name('delete_area_all');
        Route::get('/edit/{id}', 'AreaController@edit')->name('edit_area');
        Route::post('/edit/{id}', 'AreaController@postEdit')->name('post_edit_area');
        Route::post('/create', 'AreaController@create')->name('create_area');
        Route::get('/delete/{id}', 'AreaController@delete')->name('delete_area');
        Route::get('/dole', 'AreaController@doleCustomersToArea')->name('area.dole');
        Route::post('/dole', 'AreaController@postDoleCustomersToArea')->name('post.area.dole');
        Route::get('/to-user', 'AreaController@addAreaToUser')->name('add_to_user');
        Route::post('/permission/update', 'AreaController@permissionArea')->name('permission_area');
        Route::post('/move-area-back', 'AreaController@moveAreaBack')->name('move_area_back');
        Route::get('/reopen/{id}', 'AreaController@reopenArea')->name('reopen_area');
    });

    Route::group(['prefix' => 'report'], function () {
        Route::get('/', 'ReportController@index')->name('index_report');
        Route::post('/ratings', 'ReportController@ratings')->name('ratings');
    });

    Route::group(['prefix' => 'account'], function () {
        Route::get('/', 'UserController@list')->name('list_account');
        // Route::get('/create', 'UserController@create')->name('create_account');
        // Route::post('/create', 'UserController@postCreate')->name('post_create_account');
        Route::get('/edit/{id}', 'UserController@edit')->name('edit_account');
        Route::post('/edit/{id}', 'UserController@postEdit')->name('post_edit_account');
        Route::get('/delete/{id}', 'UserController@delete')->name('delete_account');
    });

    Route::group(['prefix' => 'excel'], function () {
        Route::get('/import', 'ExcelController@import')->name('data_import');
        Route::post('/import', 'ExcelController@postImport')->name('post_data_import');
        Route::get('/search', 'ExcelController@seachSHD')->name('search_shd');

        Route::get('/history', 'ExcelController@history')->name('data_import_history');
        Route::get('/customer/delete/{id}', 'ExcelController@deleteExcelCustomer')->name('delete_excel_import');

        Route::get('/appointment', 'ExcelController@appointmentExport')->name('appointment_excel');
        Route::get('/history-delete/{id}', 'ExcelController@historyDelete')->name('delete_import');


    });
    Route::group(['prefix' => 'history'], function () {
        Route::get('/area', 'HistoryController@indexArea')->name('history_area');
    });
});
