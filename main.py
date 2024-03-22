# pip install pandas
# pip install openpyxl
# ADM-Paris superpassword
# ADM-Rennes admin1password
# ADM-Strasbourg admin1password
# ADM-Grenoble admin1password
# ADM-Nantes admin1password

import pandas as pd
import random
import string
from hashlib import sha256
import getpass

try:
    ROLES = ["Admin","User"]  # Ajoutez d'autres rôles si nécessaire
    REGIONS = ["Paris", "Strasbourg", "Grenoble", "Nantes", "Rennes"]  # Mettez à jour avec vos régions
    class User:
        def __init__(self, username, password, region, role):
            self.username = username
            self.password = sha256(password.encode()).hexdigest()
            self.region = region
            self.role = role

        def check_password(self, password):
            return self.password == sha256(password.encode()).hexdigest()

    def pause():
        input("Appuyez sur une touche pour continuer...")

    def generate_password(length=8):
        characters = string.ascii_letters + string.digits + string.punctuation
        return ''.join(random.choice(characters) for i in range(length))

    def generate_username(first_name, last_name):
        return first_name[0].lower() + last_name.lower()
    
    def view_user_info(user_info):
        print("\nInformations du compte :")
        for key, value in user_info.items():
            if key != 'Password':  # Ne pas afficher le mot de passe
                print(f"{key}: {value}")

    def user_in_region(users_df, region):
        if region:
            utilisateurs_region = users_df[users_df['Region'].isin([region])]
            print(f"Utilisateurs dans la région {region}:\n{utilisateurs_region}")
        else:
            print("Tous les utilisateurs:\n", users_df)

    def setup_database(file_path):
        try:
            return pd.read_excel(file_path)
        except FileNotFoundError:
            columns = ["Nom", "Prénom", "Username", "Password", "Region", "Role", "Etat"]
            df = pd.DataFrame(columns=columns)

            # Créer un superadmin par défaut
            superadmin_data = [{
                "Nom": "Admin", 
                "Prénom": "Super", 
                "Username": "ADM-Paris", 
                "Password": sha256("superpassword".encode()).hexdigest(),  # Utiliser un mot de passe plus sûr dans un cas réel
                "Region": "Paris", 
                "Role": "SuperAdmin", 
                "Etat": "Actif"
            }]

            superadmin_df = pd.DataFrame(superadmin_data, columns=columns)
            df = pd.concat([df, superadmin_df], ignore_index=True)

            # Sauvegardez le nouveau DataFrame dans le chemin du fichier Excel spécifié
            df.to_excel(file_path, index=False)
            return df

    def add_user_to_database_xlsx(df, file_path, user, first_name, last_name, etat="Actif"):
        new_row = pd.DataFrame([{
            "Nom": last_name, 
            "Prénom": first_name, 
            "Username": user.username, 
            "Password": user.password, 
            "Region": user.region, 
            "Role": user.role,
            "Etat": etat
        }])
        df = pd.concat([df, new_row], ignore_index=True)
        df.to_excel(file_path, index=False, engine='openpyxl')
        return df
    
    def toggle_user_state(users_df, file_path, toggler_role, toggler_region=None):
        username = input("Entrez le nom d'utilisateur pour changer l'état : ")
        user = users_df[users_df['Username'] == username]

        if user.empty:
            afficher_erreur("Utilisateur non trouvé.")
            return users_df

        if toggler_role != "SuperAdmin" and user.iloc[0]['Region'] != toggler_region:
            afficher_erreur("Vous n'avez pas les droits pour modifier cet utilisateur.")
            return users_df

        new_state = select_from_menu(["Actif", "Désactivé"], "Choisissez le nouvel état pour le compte : ")
        users_df.at[user.index[0], 'Etat'] = new_state

        users_df.to_excel(file_path, index=False, engine='openpyxl')
        return users_df
    
    def login(users_df, file_path):
        username = input("Entrez votre nom d'utilisateur : ")
        attempts = 0

        while attempts < 3:
            password = password = getpass.getpass("Entrez votre mot de passe : ")
            user = users_df[(users_df['Username'] == username) & 
                            (users_df['Password'] == sha256(password.encode()).hexdigest())]

            if not user.empty:
                if user.iloc[0]['Etat'] == "Désactivé":
                    afficher_erreur("Ce compte a été désactivé.")
                    return None
                return user.iloc[0]
            else:
                afficher_erreur("Nom d'utilisateur ou mot de passe incorrect.")
                attempts += 1

        if attempts >= 3:
            afficher_erreur("Trois tentatives incorrectes. Le compte est désactivé, merci de contacter le SuperAdmin.")
            index = users_df.index[users_df['Username'] == username].tolist()
            if index:
                users_df.at[index[0], 'Etat'] = "Désactivé"
                users_df.to_excel(file_path, index=False, engine='openpyxl')
        return None

    def afficher_erreur(message):
        print("\033[91m" + message + "\033[0m")

    def afficher_information(message):
        print("\033[32m" + message + "\033[0m") 

    def Menu(message):
        print("\033[44;91m\n" + message + "\033[0m")

    def select_from_menu(options, prompt):
        for i, option in enumerate(options, 1):
            print(f"{i}. {option}")
        while True:
            try:
                choice = input(prompt)
                if choice == "":
                    return None
                choice = int(choice)
                if 1 <= choice <= len(options):
                    return options[choice - 1]
                else:
                    afficher_erreur("Choix invalide. Veuillez entrer un numéro valide.")
            except ValueError:
                afficher_erreur("Entrée invalide. Veuillez entrer un numéro.")

    def check_and_update_username(users_df, username):
        if users_df['Username'].str.contains(username).any():
            counter = 1
            new_username = f"{username}{counter}"
            while users_df['Username'].str.contains(new_username).any():
                counter += 1
                new_username = f"{username}{counter}"
            return new_username
        return username

    def create_user(users_df, file_path, creator_role, creator_region=None):
        first_name = input("Entrez le prénom de l'utilisateur : ")
        last_name = input("Entrez le nom de famille de l'utilisateur : ")
        username = generate_username(first_name, last_name)
        username = check_and_update_username(users_df, username)
        password = generate_password(random.randint(8, 12))
        afficher_information(f"Nom d'utilisateur généré : {username}\nMot de passe : {password}")

        if creator_role == "SuperAdmin":
            role = select_from_menu(ROLES, "Choisissez le rôle pour le nouveau compte : ")
            if role == "Admin":
                region = select_from_menu(REGIONS, "Choisissez la région pour le nouvel administrateur : ")
                if any((users_df['Role'] == 'Admin') & (users_df['Region'] == region)):
                    afficher_erreur(f"Un administrateur existe déjà dans la région {region}.")
                    return users_df
            else:
                region = select_from_menu(REGIONS, "Choisissez la région pour le nouveau compte : ")
        else:
            role = "User"
            region = creator_region
            afficher_information("En tant qu'administrateur, vous pouvez uniquement créer des comptes utilisateur.")

        new_user = User(username, password, region, role)
        return add_user_to_database_xlsx(users_df, file_path, new_user, first_name, last_name)

    def modify_user(users_df, file_path, modifier_role, modifier_region=None):
        username = input("Entrez le nom d'utilisateur de l'utilisateur à modifier : ")
        user = users_df[users_df['Username'] == username]

        if user.empty:
            afficher_erreur("Utilisateur non trouvé.")
            return users_df

        if modifier_role != "SuperAdmin" and user.iloc[0]['Region'] != modifier_region:
            afficher_erreur("Vous n'avez pas les droits pour modifier cet utilisateur.")
            return users_df

        while True:
            new_username = input("Entrez le nouveau nom d'utilisateur (laissez vide pour ne pas changer) : ")

            if new_username == "":
                break

            if new_username == username:
                afficher_erreur("Le nouveau nom d'utilisateur est identique à l'ancien.")
            elif (users_df['Username'] == new_username).any():
                afficher_erreur("Ce nom d'utilisateur est déjà pris. Un suffixe numérique sera ajouté.")

                counter = 1
                while (users_df['Username'] == f"{new_username}{counter}").any():
                    counter += 1

                new_username = f"{new_username}{counter}"

                afficher_information(f"Le nouveau nom d'utilisateur sera : {new_username}")

            # Déplacez cette ligne ici pour enregistrer le nouveau nom d'utilisateur avec le suffixe
            users_df.at[user.index[0], 'Username'] = new_username

            users_df.to_excel(file_path, index=False, engine='openpyxl')
            afficher_information("Nom d'utilisateur mis à jour avec succès.")
            break

        while True:
            new_password = input("Entrez le nouveau mot de passe (laissez vide pour ne pas changer) : ")
            if new_password == "":
                break  
            valid, message = is_valid_password(new_password)
            if valid:
                users_df.at[user.index[0], 'Password'] = sha256(new_password.encode()).hexdigest()
                break
            else:
                afficher_erreur(message)

        if modifier_role == "SuperAdmin":
            new_region = select_from_menu(REGIONS, "Choisissez la nouvelle région (laissez vide pour ne pas changer) : ")
            if new_region and new_region != user.iloc[0]['Region']:
                users_df.at[user.index[0], 'Region'] = new_region

        if modifier_role == "SuperAdmin":
            new_role = select_from_menu(ROLES, "Choisissez le nouveau rôle (laissez vide pour ne pas changer) : ")
            if new_role and new_role != user.iloc[0]['Role']:
                if new_role == "Admin" and any((users_df['Role'] == 'Admin') & (users_df['Region'] == user.iloc[0]['Region'])):
                    afficher_erreur(f"Un administrateur existe déjà dans la région {user.iloc[0]['Region']}.")
                else:
                    users_df.at[user.index[0], 'Role'] = new_role
                    afficher_information("Rôle de l'utilisateur mis à jour avec succès.")
        elif user.iloc[0]['Role'] != "Admin":
            afficher_erreur("Vous n'avez pas les droits pour modifier le rôle de cet utilisateur.")

        users_df.to_excel(file_path, index=False, engine='openpyxl')
        return users_df


    def delete_user(users_df, file_path, deleter_role, deleter_region=None):
        username = input("Entrez le nom d'utilisateur de l'utilisateur à supprimer : ")
        user = users_df[users_df['Username'] == username]

        if user.empty:
            afficher_erreur("Utilisateur non trouvé.")
            return users_df

        if deleter_role != "SuperAdmin" and user.iloc[0]['Region'] != deleter_region:
            afficher_erreur("Vous n'avez pas les droits pour supprimer cet utilisateur.")
            return users_df

        users_df = users_df[users_df['Username'] != username]
        users_df.to_excel(file_path, index=False, engine='openpyxl')
        return users_df
    
    def admin_menu(users_df, file_path, region):
        while True:
            Menu(f"Menu Administrateur de {region}")
            print("1. Voir les utilisateurs de cette région")
            print("2. Créer un nouvel utilisateur")
            print("3. Modifier un utilisateur")
            print("4. Supprimer un utilisateur")
            print("5. Se déconnecter")
            print("6. Activer/Désactiver un utilisateur")
            print("7. Informations du compte")

            choice = input("Entrez votre choix : ")

            if choice == "1":
                user_in_region(users_df, region)
            elif choice == "2":
                users_df = create_user(users_df, file_path, "Admin", region)
            elif choice == "3":
                users_df = modify_user(users_df, file_path, "Admin", region)
            elif choice == "4":
                users_df = delete_user(users_df, file_path, "Admin", region)
            elif choice == "5":
                break
            elif choice == "6":
                users_df = toggle_user_state(users_df, file_path, "Admin", region)
            elif choice == "7":
                view_user_info(user_info) 
            else:
                afficher_erreur("Choix invalide. Veuillez réessayer.")

    def superadmin_menu(users_df, file_path):
        while True:
            Menu("Menu Super Administrateur")
            print("1. Voir les utilisateurs (toutes les régions ou filtrer par région)")
            print("2. Créer un nouvel utilisateur")
            print("3. Modifier un utilisateur")
            print("4. Supprimer un utilisateur")
            print("5. Se déconnecter")
            print("6. Activer/Désactiver un utilisateur")
            print("7. Informations du compte")
            choix = input("Entrez votre choix : ")

            if choix == "1":
                print("0. Toutes les régions")
                for idx, region in enumerate(REGIONS, 1):
                    print(f"{idx}. {region}")
                choix_region = input("Choisissez une région ou appuyez sur 0 pour voir tous les utilisateurs : ")
                try:
                    choix_region = int(choix_region)
                    if choix_region == 0:
                        region_selectionnee = None
                    else:
                        region_selectionnee = REGIONS[choix_region - 1]
                except (ValueError, IndexError):
                    afficher_erreur("Choix invalide. Veuillez réessayer.")
                    continue
                user_in_region(users_df, region_selectionnee)
            elif choix == "2":
                users_df = create_user(users_df, file_path, "SuperAdmin")
            elif choix == "3":
                users_df = modify_user(users_df, file_path, "SuperAdmin")
            elif choix == "4":
                users_df = delete_user(users_df, file_path, "SuperAdmin")
            elif choix == "5":
                break
            elif choix == "6":
                users_df = toggle_user_state(users_df, file_path, "SuperAdmin")
            elif choix == "7":
                view_user_info(user_info) 

            else:
                afficher_erreur("Choix invalide. Veuillez réessayer.")

    def changer_nom_utilisateur(users_df, file_path, username):
        new_username = input("Entrez le nouveau nom d'utilisateur : ")
        if users_df['Username'].str.contains(new_username).any():
            afficher_erreur("Ce nom d'utilisateur est déjà pris.")
            return False

        users_df.loc[users_df['Username'] == username, 'Username'] = new_username
        users_df.to_excel(file_path, index=False, engine='openpyxl')
        print("Nom d'utilisateur mis à jour avec succès.")
        return True

        db_file = "./bdd/users.xlsx"
        users_df_xlsx = setup_database(db_file)

    def is_valid_password(password):

        if password == "":
            return False, "Le mot de passe ne doit pas être vide"
        if len(password) < 8:
            return False, "Le mot de passe doit contenir au moins 8 caractères."
        if sum(1 for c in password if c.isupper()) < 2:
            return False, "Le mot de passe doit contenir au moins 2 majuscules."
        if sum(1 for c in password if c.islower()) < 3:
            return False, "Le mot de passe doit contenir au moins 3 minuscules."
        if sum(1 for c in password if c.isdigit()) < 2:
            return False, "Le mot de passe doit contenir au moins 2 chiffres."
        if not any(c in string.punctuation for c in password):
            return False, "Le mot de passe doit contenir au moins 1 caractère spécial."
        return True, ""
    
    def changer_mot_de_passe(users_df, file_path, username):
        while True:
            new_password = input("Entrez le nouveau mot de passe : ")
            valid, message = is_valid_password(new_password)
            if valid:
                hashed_password = sha256(new_password.encode()).hexdigest()
                users_df.loc[users_df['Username'] == username, 'Password'] = hashed_password
                users_df.to_excel(file_path, index=False, engine='openpyxl')
                print("Mot de passe mis à jour avec succès.")
                break
            else:
                afficher_erreur(message)

    def menu_utilisateur(users_df, file_path, user_info):
        while True:
            nom_complet = user_info['Prénom'] + " " + user_info['Nom']
            Menu(f"Menu Utilisateur de {nom_complet}")
            print("1. Informations du compte")
            print("2. Changer le mot de passe")
            print("3. Se déconnecter")
            choix = input("Entrez votre choix : ")

            if choix == "1":
                view_user_info(user_info)
            elif choix == "2":
                changer_mot_de_passe(users_df, file_path, user_info['Username'])
            elif choix == "3":
                break
            else:
                afficher_erreur("Choix invalide. Veuillez réessayer.")


    # Initialisation des utilisateurs
    # superadmin = SuperAdmin("ADM-Paris", "superpassword", "Paris", "SuperAdmin")
    # admin1 = Admin("ADM-Rennes", "admin1password", "Rennes", "Admin")
    # admin2 = Admin("ADM-Strasbourg", "admin1password", "Starsbourg", "Admin")
    # admin3 = Admin("ADM-Grenoble", "admin1password", "Grenoble", "Admin")
    # admin4 = Admin("ADM-Nantes", "admin1password", "Nantes", "Admin")

    # # Ajout des utilisateurs à la base de données et affichage du DataFrame
    # users_df_xlsx = add_user_to_database_xlsx(users_df_xlsx, db_file, superadmin)
    # users_df_xlsx = add_user_to_database_xlsx(users_df_xlsx, db_file, admin1)
    # users_df_xlsx = add_user_to_database_xlsx(users_df_xlsx, db_file, admin2)
    # users_df_xlsx = add_user_to_database_xlsx(users_df_xlsx, db_file, admin3)
    # users_df_xlsx = add_user_to_database_xlsx(users_df_xlsx, db_file, admin4)
    db_file = "./bdd/users.xlsx"
    users_df = setup_database(db_file)


    user_info = login(users_df, db_file)
    if user_info is not None:
        role = user_info['Role']
        if role == "SuperAdmin":
            superadmin_menu(users_df, db_file)
        elif role == "Admin":
            admin_menu(users_df, db_file, user_info['Region'])
        else:
            menu_utilisateur(users_df, db_file, user_info)

except KeyboardInterrupt:
    afficher_erreur("\nInterruption du script détectée.")