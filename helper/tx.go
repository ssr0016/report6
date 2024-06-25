package helper

import (
	"database/sql"
	"fmt"
)

func CommitOrRollback(tx *sql.Tx) {
	defer func() {
		if err := recover(); err != nil {
			// Panic occurred, rollback the transaction
			if rollbackErr := tx.Rollback(); rollbackErr != nil {
				// Failed to rollback, log the error
				fmt.Printf("Failed to rollback transaction: %v\n", rollbackErr)
			}
			// Re-panic the original error
			panic(err)
		} else {
			// No panic occurred, commit the transaction
			if commitErr := tx.Commit(); commitErr != nil {
				// Failed to commit, attempt to rollback
				if rollbackErr := tx.Rollback(); rollbackErr != nil {
					// Failed to rollback after commit failure, log both errors
					fmt.Printf("Failed to commit transaction: %v\n", commitErr)
					fmt.Printf("Failed to rollback transaction after commit failure: %v\n", rollbackErr)
				} else {
					// Rolled back successfully after commit failure, log commit error
					fmt.Printf("Failed to commit transaction: %v\n", commitErr)
				}
			}
		}
	}()
}
