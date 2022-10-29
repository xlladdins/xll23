cmake_minimum_required(VERSION 3.2)
project(xll VERSION 1.0)

add_library(xll STATIC
	XLCALL.CPP
)

target_sources(xll
	XLCALL.H
	oper.h
)