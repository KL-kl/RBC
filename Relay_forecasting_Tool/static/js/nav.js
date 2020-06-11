$(function(){
    // menu收缩展开
    $('.left-menu-item>a').on('click',function(){
        if (!$('.left-menu').hasClass('left-menu-mini')) {
            if ($(this).next().css('display') == "none") {
                //展开未展开
                $('.left-menu-item').children('ul').slideUp(300);
                $(this).next('ul').slideDown(300);
                $(this).parent('li').addClass('left-menu-show').siblings('li').removeClass('left-menu-show');
            }else{
                //收缩已展开
                $(this).next('ul').slideUp(300);
                $('.left-menu-item.left-menu-show').removeClass('left-menu-show');
            }
        }
    });
    //menu-mini切换
    $('#mini').on('click',function(){
        if (!$('.left-menu').hasClass('left-menu-mini')) {
            $('.left-menu-item.left-menu-show').removeClass('left-menu-show');
            $('.left-menu-item').children('ul').removeAttr('style');
            $('.left-menu').addClass('left-menu-mini');
        }else{
            $('.left-menu').removeClass('left-menu-mini');
        }
    });
});