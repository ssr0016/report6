package pointer

func LikeString(s string) string {
	if s != "" {
		return "%" + s + "%"
	}
	return ""
}
