/*
 * Recursively expand the environment valiables, and set it.
 * On Windows XP, environment-variables expanding algorithm has a problem, variables are initialized by lexical order, there is a problem that do not expand variables recursively.
 */

var DEBUG = false;	// for debug time;

Enumerator.prototype.each = function(fn) {
	for(this.moveFirst(); !this.atEnd(); this.moveNext()) {
		fn.call(this, this.item());
	}
	return this;
}

var sh = WScript.CreateObject('WScript.Shell');
var pEnvs = new Enumerator(sh.Environment('Process'));

var allowSet = !WScript.Arguments.Named.Exists('D');
var verbose  = WScript.Arguments.Named.Exists('V');

pEnvs.each(function (item) {
	var pair = item.split('=');
	// Recursively Expanded.
	var before, after = pair[1];
	do {
		before = after;
		after  = sh.ExpandEnvironmentStrings(after);
	} while(before != after);
	if(pair[1] != after) setEnvrionmentValue(sh.Environment('Volatile'), pair[0], after, verbose, allowSet);
});

function setEnvrionmentValue(envObj, name, value, verbose, allowSet) {
	if(allowSet === (void 0)) allowSet = false;
	if(verbose || !allowSet) {
		WScript.Echo('Set ' + name + ' = ' + value);
	}
	if(allowSet) {
		envObj(name) = value;
	}
}
