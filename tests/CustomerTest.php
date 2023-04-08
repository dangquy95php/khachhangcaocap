<?php

namespace Tests;

use Maatwebsite\Excel\Facades\Excel;

class CustomerTest {
    /**
    * @test
    */
    public function user_can_download_invoices_export()
    {
        Excel::fake();

        $this->actingAs($this->givenUser())
            ->get('/invoices/download/xlsx');

        Excel::assertDownloaded('customer.xlsx', function(InvoicesExport $export) {
            // Assert that the correct export is downloaded.
            return $export->collection()->contains('#2018-01');
        });
    }
}
