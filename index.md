# Welcome to justinjtownsend.github.io
Items you find in here have largely come from my experience, but they were / are real-world challenges needing some automation / coding to resolve.

To the extent these solutions are derived from my own experiences, they are opinionated. Nevertheless, they were tested and working at the time of their use and so I share them here that they inspire the thinking of others. Code in this repo is provided 'as-is' and is not actively maintained.

- ...

## Other areas
List collections test

{% for collection in site.collections %}
  <h2>Items from {{ collection.label }}</h2>
  <ul>
    {% for item in site[collection.label] %}
      <li><a href="{{ item.url }}">{{ item.name }}</a></li>
    {% endfor %}
  </ul>
{% endfor %}
