@section('title', 'Trang chủ')
@extends('layouts.template')

@section('breadcrumb')

    <h1>DANH SÁCH KHU VỰC</h1>

    {{ Breadcrumbs::render('area') }}

@endsection

@section('content')

    <section class="section" id="area-to-user">

        <div class="row">
            <div class="col-12">
                <div id="overlay">
                    <div class="cv-spinner">
                        <span class="spinner"></span>
                    </div>
                </div>
                <div class="row">
                    <div class="col-sm-6">
                        <div class="card mb-2 full-height">
                            <div class="card-body">
                                <h5 class="card-title pb-0">Tên khu vực cần cấp quyền</h5>
                                <hr />
                                <ul class="list-group sortable flex-row" id="area_area_name">
                                    @foreach ($areas as $area)
                                        <li data-id="{{ $area->id }}"
                                            class="btn-modify d-inline-flex btn {{count($area->customers) > 0 ? 'btn-secondary' : 'btn-danger'}} pe-1 me-1 mb-1">{{ $area->name }}
                                            <span class="ms-2 badge bg-white  {{count($area->customers) > 0 ? 'text-secondary' : 'text-danger'}}">{{ count($area->customers) }}</span>
                                        </li>
                                    @endforeach
                                </ul>
                            </div>
                        </div>
                    </div>

                    <div class="col-sm-6">
                        <div id="form-user-to-area">
                            {{ csrf_field() }}
                            <div class="card mb-2">
                                <div class="card-body pe-0">
                                    <h5 class="card-title">Nhân viên cấp quyền khu vực</h5>
                                    <!-- Default Accordion -->
                                    <div class="is-scroll">
                                        @foreach ($areaUsers as $user)
                                            @if($user->username != 'admin')
                                                <div class="accordion is-item" id="accordionExample{{ $user->id }}">
                                                    <div class="accordion-item">
                                                        <h2 class="accordion-header" data-id_user="{{ $user->id }}" id="headingOne{{ $user->id }}">
                                                            <button class="accordion-button collapsed" type="button"
                                                                data-bs-toggle="collapse"
                                                                data-bs-target="#collapseOne_{{ $user->id }}"
                                                                aria-expanded="false"
                                                                aria-controls="collapseOne_{{ $user->id }}">
                                                                <b> {{ $user->username }}</b>
                                                            </button>
                                                        </h2>

                                                        <div id="collapseOne_{{ $user->id }}"
                                                            class="accordion-collapse collapse" aria-labelledby="headingOne"
                                                            data-bs-parent="#accordionExample{{ $user->id }}">
                                                            <ul class="pb-4 mb-0 is-body-user accordion-body sortable d-flex flex-wrap">
                                                                @foreach ($user->areas as $area)
                                                                    <li data-id="{{$area->id}}" class="pe-1 btn-modify btn d-inline-flex btn-secondary mb-1 me-1">
                                                                        {{ $area->name }}
                                                                        @foreach ($numberCustomerArea as $item)
                                                                            @if($area->id == $item->id)
                                                                                <span class="ms-2 badge {{ count($item->customers) > 0 ? 'bg-white text-danger' : 'bg-danger text-white'}}">{{ count($item->customers) }}</span>
                                                                            @endif
                                                                        @endforeach
                                                                    </li>
                                                                @endforeach
                                                            </ul>
                                                        </div>
                                                    </div>
                                                </div>
                                            @endif
                                        @endforeach
                                    </div>
                                </div>
                            </div>

                            <!-- <a id="btn-submit-user" class="btn btn-primary d-flex justify-content-center">Cấp Quyền</a> -->
                        </div>
                    </div>
                </div>
            </div>
        </div>
    </section>
@endsection

@push('scripts')
    <script>
        $(function() {
            var oldList, newList, item;
            $('.sortable').sortable({
                start: function(event, ui) {
                    var start_area_id  = $(ui.item).data('id');
                    var start_id_user = $(ui.item).closest('.accordion-item').children().data('id_user');
                    $(ui.item).addClass("btn-warning");

                    ui.item.data('start_area_id', start_area_id);
                    ui.item.data('start_id_user', start_id_user);

                    if ($(ui.item).closest('#area_area_name').length > 0) {
                        ui.item.data('position_area', true);
                    } else {
                        ui.item.data('position_area', false);
                    }
                },
                stop: function(event, ui) {
                    var start_id_user = ui.item.data('start_id_user');
                    var stop_id_area = ui.item.data('start_area_id');
                    var stop_id_user = $(ui.item).closest('.accordion-item').children().data('id_user');
                    $(ui.item).removeClass("btn-warning");

                    var is_move_to_left = $(ui.item).closest('#area_area_name');
                    var position_area = ui.item.data('position_area');

                    // không di chuyển qua ô bên phải
                    if (is_move_to_left.length > 0 && position_area == false) {
                        deleteArea(stop_id_area, start_id_user);
                    }

                    if (stop_id_user != undefined && (start_id_user != stop_id_user)
                    || ((start_id_user == stop_id_user) && position_area && $(ui.item).closest('.accordion-item').length > 0)
                    ) {
                        updateChangeArea(stop_id_area, stop_id_user);
                    }
                },

                connectWith: ".sortable"
            }).disableSelection().addTouch();


            function deleteArea(id_area, id_user) {
                $("#overlay").show();
                var formData = new FormData();
                formData.append('_token', "{{ csrf_token() }}");
                formData.append('area_id', id_area);
                formData.append('user_id', id_user);

                $.ajax({
                    url: "{{route('move_area_back')}}",
                    data: formData,
                    type: 'POST',
                    async: true,
                    processData: false,
                    contentType: false,
                    success:function(response) {
                        toastr.options.progressBar = true;
                        toastr.options.closeButton = true;
                        toastr.success(response.message)
                        $("#overlay").hide();
                    },
                    error: function(errors) {
                        toastr.options.progressBar = true;
                        toastr.options.closeButton = true;
                        toastr.error('Cập nhật dữ liệu thất bại.'+ errors.responseJSON.message)
                        $("#overlay").hide();
                        location.reload();
                    }
                });
            }

            function updateChangeArea(id_area, id_user) {
                $("#overlay").show();
                var formData = new FormData();
                formData.append('_token', "{{ csrf_token() }}");
                formData.append('area_id', id_area);
                formData.append('user_id', id_user);

                $.ajax({
                    url: "{{route('permission_area')}}",
                    data: formData,
                    type: 'POST',
                    async: true,
                    processData: false,
                    contentType: false,
                    success:function(response) {
                        toastr.options.progressBar = true;
                        toastr.options.closeButton = true;
                        toastr.success(response.message)
                        $("#overlay").hide();
                    },
                    error: function(errors) {
                        toastr.options.progressBar = true;
                        toastr.options.closeButton = true;
                        toastr.error('Cập nhật dữ liệu thất bại.'+ errors.responseJSON.message)
                        $("#overlay").hide();
                        location.reload();
                    }
                });
            }
        });

        ;(function($){
                $.iPhone = {
                        init: function()
                        {
                                $(window).bind('orientationchange',$.iPhone.updateOrientation);
                                this.updateOrientation();
                                $('body').css({'min-height':'420px','min-width': '320px'});
                        },
                       
                        orientation: 'portrait',
                        updateOrientation: function()
                        {
                                this.orientation = (window.orientation === 0 || window.orientation == null || window.orientation === 180) ?  'portrait' : 'landscape';
                                $('body').attr('orient',this.orientation);
                                setTimeout($.iPhone.hideURL,100);
                        },
                       
                        hideURL: function()
                        {
                                window.scrollTo(0, 1);
                                setTimeout(function(){
                                        window.scrollTo(0, 0);
                                }, 0);
                        },
                       
                        preloadImages: function(images)
                        {              
                                $(images).each(function(key,val){
                                        (new Image()).src = val;                        
                                });
                        }
                };

                $.fn.addTouch = function()
                {
                        this.each(function(i,el){
                                $(el).bind('touchstart touchmove touchend touchcancel',function(){
                                        //we pass the original event object because the jQuery event
                                        //object is normalized to w3c specs and does not provide the TouchList
                                        handleTouch(event);
                                });
                        });
                       
                        var handleTouch = function(event)
                        {
                                var touches = event.changedTouches,
                                first = touches[0],
                                type = '';
                               
                                switch(event.type)
                                {
                                        case 'touchstart':
                                                type = 'mousedown';
                                                break;
                                               
                                        case 'touchmove':
                                                type = 'mousemove';
                                                break;        
                                               
                                        case 'touchend':
                                                type = 'mouseup';
                                                break;
                                               
                                        default:
                                                return;
                                }
                               
                                var simulatedEvent = document.createEvent('MouseEvent');
                                simulatedEvent.initMouseEvent(type, true, true, window, 1, first.screenX, first.screenY, first.clientX, first.clientY, false, false, false, false, 0/*left*/, null);
                                                                                                                                         
                                first.target.dispatchEvent(simulatedEvent);
                               
                                event.preventDefault();
                        };
                };
        })(jQuery);

        $(document).ready(function(){
            $.iPhone.init();
        });

    </script>
@endpush

@push('css')
@endpush
