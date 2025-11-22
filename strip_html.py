import os.path
import bs4

def main():
    output_dir = "2025-html"
    input_dir = "2025-scrape-html"
    for f in os.listdir(input_dir):
        strip_html(os.path.join(input_dir,f), os.path.join(output_dir, f))


def strip_html(input_file, output_file):
    soup = read_soup(input_file)
    soup = strip(soup)
    write_soup(soup, output_file)


def read_soup(input_file):
    with open(input_file) as f:
        html = "".join(map(lambda x: x.strip(), f.read().splitlines()))
        soup = bs4.BeautifulSoup(html, "html.parser")
        return soup


def strip(soup):
    bad_attrs = ['data-mvc-bindings', "widgetid", 'data-dojo-attach-point', 'data-dojo-mixins', 'data-dojo-props',
                 'data-dojo-type']
    bad_tags = ["img", "script"]
    bad_elements = []

    for t in soup.self_and_descendants:
        if isinstance(t, bs4.element.Tag):
            if t.name in bad_tags:
                bad_elements.append(t)
                continue
            for a in bad_attrs:
                del t[a]

    for i in bad_elements:
        i.decompose()
    return soup


def write_soup(soup, output_file):
    with open(output_file, "w") as f:
        c = soup.prettify()
        f.write(c)

if __name__ == "__main__":
    main()