<?php

// Home
Breadcrumbs::for('home', function ($trail) {
    $trail->push('Trang chủ', route('home'));
});

Breadcrumbs::for('excel', function ($trail) {
    $trail->parent('home');
    $trail->push('Nhập dữ liệu', route('data_import'));
});

Breadcrumbs::for('list_account', function ($trail) {
    $trail->parent('home');
    $trail->push('Tài khoản nhân viên', route('list_account'));
});

Breadcrumbs::for('list_customer', function ($trail) {
    $trail->parent('home');
    $trail->push('Danh sách khách hàng', route('list_customer'));
});

Breadcrumbs::for('data_import', function ($trail) {
    $trail->parent('home');
    $trail->push('Lịc sử import', route('data_import'));
});

Breadcrumbs::for('area', function ($trail) {
    $trail->parent('home');
    $trail->push('Cấp data cho nhân viên', route('index_area'));
});

Breadcrumbs::for('area.dole', function ($trail) {
    $trail->parent('home');
    $trail->push('Chia đều data', route('area.dole'));
});

Breadcrumbs::for('report', function ($trail) {
    $trail->parent('home');
    $trail->push('Báo cáo cuộc gọi nhân viên', route('index_report'));
});

Breadcrumbs::for('statistical', function ($trail) {
    $trail->parent('home');
    $trail->push('Thống kê khách liệu khách hàng của nhân viên', route('index_statistical'));
});

Breadcrumbs::for('history_area', function ($trail) {
    $trail->parent('home');
    $trail->push('Lịch sử khu vực', route('history_area'));
});

