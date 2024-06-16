#include <node.h>
#include <string>
#include "xlsx.h"
#include <fstream>
#include <iostream>

using v8::Context;
using v8::FunctionCallbackInfo;
using v8::Isolate;
using v8::Local;
using v8::Object;
using v8::String;
using v8::Number;
using v8::Value;
using v8::Exception;

void setValue(const FunctionCallbackInfo<Value>& args) {
	Isolate* isolate = args.GetIsolate();
	Local<Context> context = isolate->GetCurrentContext();
	Local<Object> self = args.This();
	if (args.Length() != 2 || !args[0]->IsString()) {
		isolate->ThrowException(Exception::TypeError(String::NewFromUtf8(isolate, "Wrong arguments").ToLocalChecked()));
		return;
	}

	Local<String> extName = String::NewFromUtf8(isolate, "_sheet_").ToLocalChecked();
	Local<v8::External> ext = self->Get(context, extName).ToLocalChecked().As<v8::External>();
	document::Sheet* sheet = (document::Sheet*)ext->Value();

	String::Utf8Value utf8(isolate, args[0]);
	std::string x = std::string(*utf8);
	std::fstream f;

	Local<Value> value = args[1];
	if (value->IsString()) {
		Local<String> local = args[1].As<v8::String>();
		String::Utf8Value utf8(isolate, local);
		std::string val = std::string(*utf8, utf8.length());
		sheet->put_str(x, val);
	}
	if (value->IsNumber()) {
		double val = value.As<Number>()->Value();
		sheet->put_dbl(x, val);
	}

}

void newSheet(const FunctionCallbackInfo<Value>& args) {
	Isolate* isolate = args.GetIsolate();
	Local<Context> context = isolate->GetCurrentContext();
	if (args.Length() != 1 || !args[0]->IsString()) {
		isolate->ThrowException(Exception::TypeError(String::NewFromUtf8(isolate, "Wrong arguments").ToLocalChecked()));
		return;
	}
	Local<v8::Value> arg = args[0];
	String::Utf8Value utf8(isolate, arg);
	std::string name = std::string(*utf8);


	Local<Object> self = args.This();
	Local<String> __doc = String::NewFromUtf8(isolate, "_doc_").ToLocalChecked();
	Local<v8::External> _doc = self->Get(context, __doc).ToLocalChecked().As<v8::External>();
	document* doc = (document*)_doc->Value();

	document::Sheet* sheet = doc->new_sheet(name);

	Local<Object> obj = Object::New(isolate);

	obj->Set(
		context,
		String::NewFromUtf8(isolate, "set").ToLocalChecked(),
		v8::FunctionTemplate::New(isolate, setValue)->GetFunction(context).ToLocalChecked()
	);

	obj->Set(
		context,
		String::NewFromUtf8(isolate, "docName").ToLocalChecked(),
		String::NewFromUtf8(isolate, doc->getName().c_str()).ToLocalChecked()
	).FromJust();

	obj->Set(
		context,
		String::NewFromUtf8(isolate, "sheetName").ToLocalChecked(),
		String::NewFromUtf8(isolate, name.c_str()).ToLocalChecked()
	).FromJust();

	obj->Set(
		context,
		String::NewFromUtf8(isolate, "_doc_").ToLocalChecked(),
		v8::External::New(isolate, doc)
	).FromJust();

	obj->Set(
		context,
		String::NewFromUtf8(isolate, "_sheet_").ToLocalChecked(),
		v8::External::New(isolate, sheet)
	).FromJust();

	args.GetReturnValue().Set(obj);
}

void getSheet(const FunctionCallbackInfo<Value>& args) {
	Isolate* isolate = args.GetIsolate();
	Local<Context> context = isolate->GetCurrentContext();
	if (args.Length() != 1 || !args[0]->IsString()) {
		isolate->ThrowException(Exception::TypeError(String::NewFromUtf8(isolate, "Wrong arguments").ToLocalChecked()));
		return;
	}
	Local<v8::Value> arg = args[0];
	String::Utf8Value utf8(isolate, arg);
	std::string name = std::string(*utf8);


	Local<Object> self = args.This();
	Local<String> __doc = String::NewFromUtf8(isolate, "_doc_").ToLocalChecked();
	Local<v8::External> _doc = self->Get(context, __doc).ToLocalChecked().As<v8::External>();
	document* doc = (document*)_doc->Value();

	document::Sheet* sheet = doc->get_sheet(name);
	if (sheet == nullptr) {
		isolate->ThrowException(Exception::TypeError(String::NewFromUtf8(isolate, "There is no such sheet").ToLocalChecked()));
	}

	Local<Object> obj = Object::New(isolate);

	obj->Set(
		context,
		String::NewFromUtf8(isolate, "set").ToLocalChecked(),
		v8::FunctionTemplate::New(isolate, setValue)->GetFunction(context).ToLocalChecked()
	);

	obj->Set(
		context,
		String::NewFromUtf8(isolate, "docName").ToLocalChecked(),
		String::NewFromUtf8(isolate, doc->getName().c_str()).ToLocalChecked()
	).FromJust();

	obj->Set(
		context,
		String::NewFromUtf8(isolate, "sheetName").ToLocalChecked(),
		String::NewFromUtf8(isolate, name.c_str()).ToLocalChecked()
	).FromJust();

	obj->Set(
		context,
		String::NewFromUtf8(isolate, "_doc_").ToLocalChecked(),
		v8::External::New(isolate, doc)
	).FromJust();

	obj->Set(
		context,
		String::NewFromUtf8(isolate, "_sheet_").ToLocalChecked(),
		v8::External::New(isolate, sheet)
	).FromJust();

	args.GetReturnValue().Set(obj);
}

void save(const FunctionCallbackInfo<Value>& args) {
	Isolate* isolate = args.GetIsolate();
	Local<Context> context = isolate->GetCurrentContext();
	if (args.Length() != 0) {
		isolate->ThrowException(Exception::TypeError(String::NewFromUtf8(isolate, "Wrong arguments").ToLocalChecked()));
		return;
	}
	Local<Object> self = args.This();
	Local<String> __doc = String::NewFromUtf8(isolate, "_doc_").ToLocalChecked();
	Local<v8::External> _doc = self->Get(context, __doc).ToLocalChecked().As<v8::External>();
	document* doc = (document*)_doc->Value();

	doc->save();
}

void newDocument(const FunctionCallbackInfo<Value>& args) {
	Isolate* isolate = args.GetIsolate();
	Local<Context> context = isolate->GetCurrentContext();
	if (args.Length() != 1 || !args[0]->IsString()) {
		isolate->ThrowException(Exception::TypeError(String::NewFromUtf8(isolate, "Wrong arguments").ToLocalChecked()));
		return;
	}
	Local<String> local = args[0].As<v8::String>();
	String::Utf8Value utf8(isolate, local);
	std::string name = std::string(*utf8, utf8.length());

	Local<Object> obj = Object::New(isolate);
	document* doc = new document(name);

	obj->Set(
		context,
		String::NewFromUtf8(isolate, "docName").ToLocalChecked(),
		String::NewFromUtf8(isolate, name.c_str()).ToLocalChecked()
	).FromJust();

	obj->Set(
		context,
		String::NewFromUtf8(isolate, "_doc_").ToLocalChecked(),
		v8::External::New(isolate, doc)
	).FromJust();

	obj->Set(
		context,
		String::NewFromUtf8(isolate, "newSheet").ToLocalChecked(),
		v8::FunctionTemplate::New(isolate, newSheet)->GetFunction(context).ToLocalChecked()
	);

	obj->Set(
		context,
		String::NewFromUtf8(isolate, "getSheet").ToLocalChecked(),
		v8::FunctionTemplate::New(isolate, getSheet)->GetFunction(context).ToLocalChecked()
	);

	obj->Set(
		context,
		String::NewFromUtf8(isolate, "save").ToLocalChecked(),
		v8::FunctionTemplate::New(isolate, save)->GetFunction(context).ToLocalChecked()
	);


	args.GetReturnValue().Set(obj);
}

void init(Local<Object> exports) {
	NODE_SET_METHOD(exports, "newDocument", newDocument);
}

NODE_MODULE(NODE_GYP_MODULE_NAME, init)