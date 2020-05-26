(function($) { 

var vInt=0; // this variable controls the loop
var refresh=1; // refresh when a time finish
var interval=1000; // the loop interval


$.jheartbeat = {

    options: {
		url: "",
		delay: 10000,
		div_id: "test_div"
    },
	
	beatfunction:  function(){
	
	},
	
	timeoutobj:  {
		id: -1
	},

    set: function(options, onbeatfunction) {
		if (this.timeoutobj.id > -1) {
			clearTimeout(this.timeoutobj);
		}
        if (options) {
            $.extend(this.options, options);
        }
        if (onbeatfunction) {
            this.beatfunction = onbeatfunction;
        }

		// Add the HeartBeatDIV to the page
		$("body").append("<div id=\"" + this.options.div_id + "\" style=\"display: none;\"></div>");
		this.timeoutobj.id = setTimeout("$.jheartbeat.beat();", this.options.delay);
    },

    beat: function() {
		$.ajax({
				url: this.options.url,
				dataType: "html",
				type: "GET",
				error: function(e)   { 
					$('#'+ $.jheartbeat.options.div_id).append("Error Requesting Data"); 
				},
				success: function(data){ 
					$('#'+ $.jheartbeat.options.div_id).html(data); 
				}
			   });
		this.timeoutobj.id = setTimeout("$.jheartbeat.beat();", this.options.delay);
        this.beatfunction();
    }
};


// this function autostarts the infinite loop, every second, triggers the countdown fn
jQuery.autocountdown = function () {
	$('.countdown').countdown(); // trigger the fn
	vInt=setInterval("$('.countdown').countdown();", interval); // set the loop
}

// countdown function, update second-by-second the time to finish
jQuery.fn.countdown = function (options) {
	var defaults = {  // set up default options
		refresh:     1,		 // refresh when a time finish
		interval:    1000, // the loop interval
		cdClass:     'countdown', // the class to apply this plugin
		granularity: 3,
		
		label:    ['s ', 'd ', 'h', ':', ''],
		units:    [604800, 86400, 3600, 60, 1]
	};
	if (options && options.label) {
		$.extend(defaults.label, options.label);
		delete options.label;
	}
    if (options && options.units) {
      $.extend(defaults.units, options.units);
      delete options.units;
    }
	$.extend(defaults, options);

	// pad fn, add left zeros to the string
	var pad = function (value, length) {
		value = String(value);
		length = parseInt(length) || 2;
		while (value.length < length)
			value = "0" + value;
		if (value<1) value = "00";
		return value;
	};

	var format_interval = function (timestamp) {
		var label = defaults.label;
		var units = defaults.units;
		var granularity = defaults.granularity;

		output = '';
		for (i=1; i<=units.length; i++) {
			value=units[i];
			if (timestamp >= value) {	      				
				var val=pad(Math.floor(timestamp / value), 2);
				val = val>0 ? val : '00';
				output += val + label[i];
				timestamp %= value;
				granularity--;
			} 
			else if (value==1) output += '00'; // we need the final seconds to allways show 00, i.e., 03:00

			if (granularity == 0)
				break; 
		}
		
		if (output.length<3) output = '00:'+output;
		return output ? output : '00:00';
	}
	
	// the countdown core
	return this.each(function() {
		secs=$(this).attr('secs');
		$(this).html(format_interval(secs));
		secs--;
		
		if (secs<1) {
			$(this).attr('secs', '...');
			clearInterval(vInt);
			if (refresh)
				window.location.href = window.location.href;
		} else
			$(this).attr('secs', secs);
		
	});

}

	$.autocountdown(); // loop
})(jQuery);