from datetime import date
from sys import exit as sys_exit


def validate_config(LOGGER, data: dict):
    root = ["courses", "semester", "term"]
    term_info = ["start", "end"]

    # checking config root keys
    missing_keys = set(root) - set(data.keys())
    if missing_keys:
        LOGGER.error("Missing root key: %s", ", ".join(missing_keys))
        sys_exit(0)

    # checking semester name
    if data["semester"].get("name"):
        if not isinstance(data["semester"]["name"], str):
            LOGGER.error(
                "expected semester name to be str but got %s",
                type(data["semester"]["name"]),
            )
            sys_exit(0)
    else:
        LOGGER.error("missing semester name")
        sys_exit(0)

    # checking term values
    for k, v in data["term"].items():
        missing_keys = set(term_info) - set(v.keys())
        if missing_keys:
            LOGGER.error("Missing %s term key: %s", k, ", ".join(missing_keys))
            sys_exit(0)
        else:
            for k1, v1 in v.items():
                if date != v1.__class__:
                    LOGGER.error(
                        "%s.%s requires <class 'datetime.date'> but got %s",
                        k,
                        k1,
                        type(v1),
                    )
                    sys_exit(0)

    # checking courses values
    for course in data["courses"]:
        c_data = data["courses"][course]

        if not c_data.get("name"):
            LOGGER.error("%s - missing course name", course)
            sys_exit(0)
        else:
            if not isinstance(c_data["name"], str):
                LOGGER.error(
                    "%s - expected course name to be str but got %s",
                    course,
                    type(c_data["name"]),
                )
                sys_exit(0)

        if not c_data.get("days"):
            LOGGER.error("%s - missing days", course)
            sys_exit(0)
        else:
            if not isinstance(c_data["days"], list):
                LOGGER.error(
                    "%s - expected days to be an array but got %s",
                    course,
                    type(c_data["days"]),
                )
                sys_exit(0)
