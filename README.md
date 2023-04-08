
Hi em!
Anh xin phép gửi 1 số chức năng trong giai đoạn 1 để bàn giao bên em 

Lịch trình bắt đầu tiến hành viết phần mềm(Gia đoạn 1) 3/6/2022 - 9/6/2022
Tài khoản nhân viên:
- Đăng ký sản tài khoản
- Danh sách tài khoản
Thông tin file data
- Xử lý dữ liệu sai sót và trùng lập trước khi import
- Xử lý dữ liệu và đưa số dòng dữ liệu cho nhân viên tuỳ ý
Xử lý dữ liệu khách hàng
- Nhật ký cuộc gọi
- Danh sách cuộc gọi
- Tạo data giả số tiền cho  nhân viên
- Cấp quyền nhân viên theo khu vực
- Lưu dữ liệu khi sau khi thao tác với khách hàng

Thời gian bàn giao sản phẩm giai đoạn 1: 10/6/2022 - 11/6/2022

GRANT CREATE, ALTER, DROP, INSERT, UPDATE, DELETE, SELECT, REFERENCES, RELOAD on khachangchatluong1.* TO 'khachhangchatluong'@'localhost' WITH GRANT OPTION;

### Hiển thị databases;
`show databases`

### Show user trong mysql
`SELECT user FROM mysql.user;`

### Tạo user
`CREATE USER 'newuser'@'localhost' IDENTIFIED BY 'password';`
### Xoá user
`DROP USER 'newuser'@'localhost';`

### Tạo databases
`create database db_test`

### Cấp quyền user cho database
`GRANT ALL ON *.* TO 'user'@'localhost';`

### Chạy lệnh FLUSH 
`FLUSH PRIVILEGES;`

### Restart the mysqld service:
`sudo service mysqld restart`

### Confirm that the server has successfully restarted:
`sudo service mysqld status`

### Thay đổi time-zone mysql trên server UTC+07:00
`sudo nano /etc/mysql/my.cnf`
`[mysqld]
default-time-zone = "+07:00"`
`sudo service mysql restart`

### Kiểm tra time-zone hiện tại
`sudo mysql –e "SELECT @@global.time_zone;"`
`sudo mysql –e "SELECT NOW();"`


### Mở thêm thời gian tải cho apache
`max_execution_time = 360      ; Maximum execution time of each script, in seconds (I CHANGED THIS VALUE)
max_input_time = 120          ; Maximum amount of time each script may spend parsing request data
;max_input_nesting_level = 64 ; Maximum input variable nesting level
memory_limit = 128M           ; Maximum amount of memory a script may consume (128MB by default)
`

### Xóa một số cache và config cache
`php artisan view:clear
php artisan view:cache
php artisan route:clear
php artisan route:cache
php artisan config:clear
php artisan config:cache`

### Run migrate
`php artisan migrate --path="/database/migrations/2022_08_18_153044_create_history_excel_table.php"`

### Config source code khachhangchatluong
`DATE_REOPEN = 2022-06-10
SLACK_REPORT_ENABLED=true
LOG_SLACK_WEBHOOK_URL=https://hooks.slack.com/services/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`
