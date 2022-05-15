# Welcome to justinjtownsend.github.io
Items on here have come from situations in my own work, but they are real-world challenges needing a coding approach to resolve. This set of pages takes a pracitioner perspectives because... *something from George Bernard Shaw*

Technical skills are increasingly important for productivity and reliance on others is not always possible. This is especially true in lean (small) organisations (hedge funds, startups, charities, etc.).

To the extent these solutions are derived from my own experiences, they are opinionated. Nevertheless, they were tested and working at the time of their use and so I share them here that they inspire the thinking of others. Code in this repo is provided 'as-is' and is not actively maintained.

Thank you to GitHub, Jekyll, Ruby and Liquid with you I could not keep on top of this stuff. You're awesome. :100:

## Collections

{% for collection in site.collections %}
  <h2>{{ collection.title }}</h2>
  <ul>
    {% for item in site[collection.label] %}
      <li><a href="{{ item.url }}">{{ item.title }}</a></li>
    {% endfor %}
  </ul>
{% endfor %}
