function notify(text, type) {
    new Noty({
        text: text,
        theme: 'mint',
        type: type,
        timeout: type=='success'?1000:5000,
        progressBar: true,
        closeWith: ['click'],
    }).show();
}

document.onkeydown = function(e){
    
    if(e.key == 'Enter'){
        document.querySelector('.button.submit').click();
    }
}