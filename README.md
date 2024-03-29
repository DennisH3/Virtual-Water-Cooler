[(Français)](#le-nom-du-projet)

# Virtual Water Cooler

The Virtual Water Cooler (VWC) aims to mitigate some of the disconnect many feel from remote work by having employees sign up through a Google Forms (bilingual) and then automatically being paired and notified through email. 
It is a two-part program: pair matching and email automation.

The code can be modified to create pairs from different forms and different requirements or form groups larger than pairs.

## Files

### Main program
- Virtual Water Cooler v2.py: merges the English and French responses and creates pairs based on time and language preferences. Additionally, it outputs a list of the pairs' common interests.
- VWCemail.py: automates sending emails per matched pair

### Translation
- google_trans_new.py: the package that translates French responses to English, and English results to French. It is included here as the most recent version which fixes a bug has not been published to PyPi
- constant.py: dependency for google_trans_new.py

### How to Contribute

See [CONTRIBUTING.md](CONTRIBUTING.md)

### License

Unless otherwise noted, the source code of this project is covered under Crown Copyright, Government of Canada, and is distributed under the [MIT License](LICENSE).

The Canada wordmark and related graphics associated with this distribution are protected under trademark law and copyright law. No permission is granted to use them outside the parameters of the Government of Canada's corporate identity program. For more information, see [Federal identity requirements](https://www.canada.ca/en/treasury-board-secretariat/topics/government-communications/federal-identity-requirements.html).

______________________

## Le nom du projet

- Quel est ce projet?
- Comment ça marche?
- Qui utilisera ce projet?
- Quel est le but de ce projet?

### Comment contribuer

Voir [CONTRIBUTING.md](CONTRIBUTING.md)

### Licence

Sauf indication contraire, le code source de ce projet est protégé par le droit d'auteur de la Couronne du gouvernement du Canada et distribué sous la [licence MIT](LICENSE).

Le mot-symbole « Canada » et les éléments graphiques connexes liés à cette distribution sont protégés en vertu des lois portant sur les marques de commerce et le droit d'auteur. Aucune autorisation n'est accordée pour leur utilisation à l'extérieur des paramètres du programme de coordination de l'image de marque du gouvernement du Canada. Pour obtenir davantage de renseignements à ce sujet, veuillez consulter les [Exigences pour l'image de marque](https://www.canada.ca/fr/secretariat-conseil-tresor/sujets/communications-gouvernementales/exigences-image-marque.html).
