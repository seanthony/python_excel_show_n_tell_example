from json import load


def load_settings(filepath):
    with open(filepath) as file:
        return load(file)


def main():
    settings = load_settings("./settings.json")


if __name__ == "__main__":
    main()
