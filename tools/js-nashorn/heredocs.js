var data = {
	foo: "bar",
	time: new Date()
};

print(<<EOF);
So...
	foo = ${data.foo}
and the current time is
	${data.time}
EOF
