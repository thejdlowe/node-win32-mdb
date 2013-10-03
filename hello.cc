#include <node.h>
#include <v8.h>
#include <iostream> 
#include <string> 
 
#import "C:\Program files\Common Files\System\Ado\msado15.dll" rename("EOF", "ADOEOF")

using namespace v8;

Handle<Value> Method(const Arguments& args) {
	HandleScope scope;
	HRESULT hr;
	CoInitialize(NULL);
	try {
		ADODB::_ConnectionPtr myConn;
		hr = myConn.CreateInstance(__uuidof(ADODB::Connection));
		if(FAILED(hr)) {
			throw _com_error(hr);
		}
	}
	catch(_com_error &e) {
		printf("Error:\n");
		printf("Code = %08lx\n", e.Error());
		printf("Message = %s\n", e.ErrorMessage());
		printf("Source = %s\n", (LPCSTR) e.Source());
		printf("Description = %s\n", (LPCSTR) e.Description());
	}
	/*ADODB::_RecordsetPtr myRecord; 
	hr = myRecord.CreateInstance(__uuidof(ADODB::Recordset)); 
	if (FAILED(hr)) { 
		throw _com_error(hr); 
	}*/

	if(!args[0]->IsString()) {
		//ThrowException(Exception::Error(String::New("FAIL!")));
		return scope.Close(String::New("FAIL"));
	}
	else return scope.Close(String::New("world"));
}

void init(Handle<Object> exports) {
	exports->Set(String::NewSymbol("hello"),
		FunctionTemplate::New(Method)->GetFunction());
}

NODE_MODULE(hello, init)