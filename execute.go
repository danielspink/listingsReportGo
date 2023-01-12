package main

import (
	"fmt"
	"os/exec"
)

func runCom() {
	cmd := exec.Command("powershell", "ls", exportDirectory)

	out, err := cmd.Output()
	if err != nil {
		fmt.Println("Did not find directory: ", exportDirectory)
	} else {
		fmt.Println(string(out))
	}
}
