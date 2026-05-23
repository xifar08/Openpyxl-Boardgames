from utils import get_data
from config import URL


def main():
    df = get_data(URL, sep=';')
    print(df.head(5))


if __name__ == "__main__":
    main()
