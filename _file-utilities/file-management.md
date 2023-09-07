---
layout: default
title: File Management
topic: File Management / Manipulation
---

## File Management
Organisation, tagging and converting of files remains a non-trivial task in modern computing.

Executives rightly continue to have a preference for communicating their message with the most appropriate medium, so converting information efficiently between file formats for various purposes is often a handy skill:

- in this case executive-level presentations, non-editability (of the content) was an important requirement for preservation of ground truth. Below are a couple of approaches using the popular Adobe Acrobat engine and (at the time of development) the more cost-effective PowerPDF from Nuance Communications. Arguably such approaches have limited application where modern image-scraping techniques bypass 'non-editable' content...

  - [convert-pdf-acrobat.ps1]({% link /assets/samples/convert-pdf-acrobat.ps1 %})
  - [convert-pdf-nuance.ps1]({% link /assets/samples/convert-pdf-nuance.ps1 %})

- a symptom of the proliferation of digital cameras and image formats is the varied means of transfer and naming conventions, this script represents an (incomplete) approach to organising images based on the date the photo was shot and takes advantage of [Hugsan's EXIFUtils](http://www.hugsan.com/EXIFutils/)
  - [organize-photos-exifutils.ps1]({% link /assets/samples/organize-photos-exifutils.ps1 %})
