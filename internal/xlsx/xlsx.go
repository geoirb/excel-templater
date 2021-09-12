package xlsx

import (
	"fmt"
)

func getColByIdx(colIdx int) string {
	return fmt.Sprintf("%c", 'A'+colIdx)
}
