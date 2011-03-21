all: avg_counter
obj = main.o calc.o handler.o
avg_counter: $(obj) 
	g++ $(obj) -lexcelformat -o avg_counter
main.o: main.cpp
	g++ -c main.cpp
calc.o: calc.cpp calc.h
	g++ -c calc.cpp
handler.o: handler.h handler.cpp
	g++ -c handler.cpp
clean: 
	rm *.o

