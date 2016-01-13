
// tbMask v1.0 
// 
// Textbox 'masking' on client-side in Internet Explorer. 
// 
// Written By John McGlothlin - April 17th, 2004 
// 
// Mask Characters 
// 9 = Numeric only 
// A = Alpha only 
// X = Alphanumeric 
// 
// All other chars treated as literals. 
// 
// 
var CTRL_PASTE = 22; 
var CTRL_COPY = 3; 
var CTRL_CUT = 24; 
var TAB_KEY = 9; 
var DELETE_KEY = 46; 
var BACKSPACE_KEY = 8; 
var ENTER_KEY = 13; 
var RIGHT_ARROW_KEY = 39; 
var DOWN_ARROW_KEY = 40; 
var UP_ARROW_KEY = 38; 
var LEFT_ARROW_KEY = 37; 
var HOME_KEY = 36; 
var END_KEY = 35; 
var PAGEUP_KEY = 33; 
var PAGEDOWN_KEY = 34; 
var CAPS_LOCK_KEY = 20; 
var ESCAPE_KEY = 27; 
// -------------------------------------------------------------------- 
// Main function 
// Gets the current keystroke and deals with it....lol 
// -------------------------------------------------------------------- 
function tbMask(textBox){ 
	var keyCode = event.keyCode; 
// get character from keyCode....dealing with the "Numeric KeyPad" 
// keyCodes so that it can be used 
	var keyCharacter = cleanKeyCode(keyCode); 
	var retVal = false; 
// grab the mask 
	var mask = textBox.mask; 
	switch(keyCode){ 
		case BACKSPACE_KEY: 
			var c = getCursorPos(textBox); 
			if(c > 0){ 
				var currentMaskChar; 
// get next available char to delete except mask chars 
				while(c > 0){ 
					c--; 
					currentMaskChar = mask.charAt(c); 
					if(currentMaskChar == '9' || currentMaskChar == 'X' || currentMaskChar == 'A'){ 
// found a spot.....replace that char with '_' 
						var x = textBox.value.substring(0,c); 
						var y = textBox.value.substring(c+1,textBox.value.length); 
						textBox.value = x + '_' + y;
						setCursorPos(textBox,c);
						textBox.curPos = c;
						break; 
					} 
				} 
			}
			break; 

		case TAB_KEY: // keep track of cursor b4 tabbing out of field 
			var c = getCursorPos(textBox); 
			textBox.curPos = c; 
			retVal = true; 
			break; 

		case HOME_KEY: // just move/keep track of cursor 
			setCursorPos(textBox,0); 
			textBox.curPos = c; 
			break; 
		
		case END_KEY: // just move/keep track of cursor 
			setCursorPos(textBox,textBox.value.length); 
			textBox.curPos = textBox.value.length; 
			break; 
		
		case ENTER_KEY:
			retVal = true; 
			break; 

		case DELETE_KEY:
			var c = getCursorPos(textBox); 
			if(c > -1){ 
				var currentMaskChar = mask.charAt(c); 
// only allow delete if it's a valid char 
				if(currentMaskChar == '9' || currentMaskChar == 'X' || currentMaskChar == 'A'){ 
					var x = textBox.value.substring(0,c); 
					var y = textBox.value.substring(c+1,textBox.value.length); 
					textBox.value = x + '_' + y; 
					setCursorPos(textBox,c); 
					textBox.curPos = c; 
				} 
			} 
			break; 
		
		case LEFT_ARROW_KEY: // just move/keep track of cursor 
			var c = getCursorPos(textBox); 
			if(c > 0){ 
				setCursorPos(textBox,c-1); 
				textBox.curPos = c-1; 
			} 
			break; 
			
		case RIGHT_ARROW_KEY: // just move/keep track of cursor 
			var c = getCursorPos(textBox); 
			if(c < textBox.value.length) { 
				setCursorPos(textBox,c+1); 
				textBox.curPos = c+1; 
			} 
			break; 
			
		default: // adding a new char somewhere in the field 
			var c = getCursorPos(textBox); 
			var currentMaskChar; 
// get next available to change.....except masking chars 
			while(c < textBox.value.length){ 
				currentMaskChar = mask.charAt(c); 
				if(currentMaskChar == '9' || currentMaskChar == 'X' || currentMaskChar == 'A') break; 
				c++; 
			} 
			switch(currentMaskChar){ 
				case '9': // numeric only 
					if('0123456789'.indexOf(keyCharacter) != -1) 
						addNewKey(textBox,keyCharacter,c); 
					break; 
				
				case 'A': // alpha only 
					if('abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ'.indexOf(keyCharacter) != -1) 
					addNewKey(textBox,keyCharacter,c); 
					break; 
				
				case 'X': // alphanumeric 
					if('abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789'.indexOf(keyCharacter) != -1) 
					addNewKey(textBox,keyCharacter,c); 
					break; 
					default: 
			break; 
		} 
	} 
	return retVal; 
} 

// -------------------------------------------------------------------- 
// Accept a TextBox, a keycharacter and an integer position 
// adds the key to the position and then sets cursor to next availble spot 
// -------------------------------------------------------------------- 
function addNewKey(tb,key,pos){ 
// add the new key to the textbox at pos 
	var startSel = tb.value.substring(0,pos); 
	var endSel = tb.value.substring(pos+1,tb.value.length); 
	tb.value = startSel + key + endSel; 
// advance cursor to next '_' 
	while(pos < tb.value.length){ 
		curChar = tb.value.charAt(pos); 
		if(curChar == '_') break; 
		pos++; 
	} 
	setCursorPos(tb,pos); 
	tb.curPos = pos; 
} 

// -------------------------------------------------------------------- 
// Loops thru pasted value and checks each char against the mask 
// 
// Leaves old value and return false if *any* char is off 
// Creates new masked value if all pasted data is ok 
// -------------------------------------------------------------------- 

function tbPaste(textBox) { 
// grab the textBox value and the mask 
	var pastedVal = window.clipboardData.getData("Text"); 
	var mask = textBox.mask; 
	var newVal = ''; 
	var curPastedVal = 0;
	for(var i=0;i<mask.length;i++){ 
		var currentMaskChar = mask.charAt(i); 
// if current mask pos allows entry 
		if(currentMaskChar == '9' || currentMaskChar == 'X' || currentMaskChar == 'A'){ 
			var currentPastedChar = pastedVal.charAt(curPastedVal); 
// check each current mask char against new keystroke 
// return false if any are out of sync 
			if(currentMaskChar == '9'){ 
				if('0123456789'.indexOf(currentPastedChar) == -1) 
				return false; 
			} else if(currentMaskChar == 'A') { 
				if('abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ'.indexOf(currentPastedChar) == -1) 
					return false; 
			} else if(currentMaskChar == 'X') { 
				if('abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789'.indexOf(currentPastedChar) == -1) 
					return false; 
			} else { 
				return false; 
			} 
// add new key 
			newVal += currentPastedChar; 
			curPastedVal++; 
		}else{ 
// add mask literal 
			newVal += currentMaskChar; 
		} 
	} 
	textBox.value = newVal; 
	return false; 
} 

// -------------------------------------------------------------------- 
// puts masked thingie in tb ie: '(___) ___-____' 
// -------------------------------------------------------------------- 
function tbFocus(textBox) { 
	var val = textBox.value; 
	var mask = textBox.mask; 
	var startVal = ''; 
// give the tb a property set to current cursor pos...sorta used at this point 
// to set cursor pos when re-entering a tb. 
// eventually want to use it to get rid of get/setCursorPos functions...if possible 
	if(textBox.curPos == 'undefined' || textBox.curPos == null) 
		textBox.curPos = -1; 
	// if value null....set the starting format ie: '(___) ___-____' 
	if(val.length == 0 || val == null) { 
		for(var i = 0; i < mask.length; i++) { 
			var c = mask.charAt(i); 
			if(c == '9' || c == 'X' || c == 'A') { 
				startVal += '_'; 
				if(textBox.curPos == -1) textBox.curPos = i; 
			}else{ 
				startVal += c; 
			} 
		} 
	textBox.value = startVal; 
	} 
// otherwise just set proper cursor pos 
	if(textBox.curPos == -1) { 
		textBox.curPos = textBox.value.length; 
	} 
	setCursorPos(textBox,textBox.curPos); 
// set just in case. 
	textBox.maxlength = mask.length; 
	return true; 
} 

// -------------------------------------------------------------------- 
// The Numeric KeyPad returns keyCodes that ain't all that workable. 
// 
// ie: KeyPad '1' returns keyCode 97 which String.fromCharCode converts to an 'a'. 
// 
// This way allows the Numeric KeyPad to be used 
// -------------------------------------------------------------------- 

function cleanKeyCode(key){ 
	switch(key){ 
		case 96: return "0"; break; 
		case 97: return "1"; break; 
		case 98: return "2"; break; 
		case 99: return "3"; break; 
		case 100: return "4"; break; 
		case 101: return "5"; break; 
		case 102: return "6"; break; 
		case 103: return "7"; break; 
		case 104: return "8"; break; 
		case 105: return "9"; break; 
		default: return String.fromCharCode(key); break; 
	} 
} 

// -------------------------------------------------------------------- 
// Google gems 
// -------------------------------------------------------------------- 

function getCursorPos(el){ 
	var sel, rng, r2, i=-1; 
	if(document.selection && el.createTextRange) { 
		sel=document.selection; 
		if(sel){ 
			r2=sel.createRange(); 
			rng=el.createTextRange(); 
			rng.setEndPoint("EndToStart", r2); 
			i=rng.text.length; 
		} 
	} 
	return i; 
} 

function setCursorPos(field,pos) { 
	if (field.createTextRange) { 
		var r = field.createTextRange(); 
		r.moveStart('character', pos); 
		r.collapse(); 
		r.select(); 
	} 
} 

