#!/bin/bash

echo " Welcome, plese select your class:
1 - Samurai
2 - Prisoner
3 - Prophet"

read class

case $class in

        1)
                type="Samurai"
                hp=10
                attack=20
                ;;
        2)
                type="Prisoner"
                hp=20
                attack=4
                ;;
        3)
                type="Prophet"
                hp=30
                attack=4
                ;;
esac

echo "your class is $type your hp is $hp, and your attack is $attack"

echo "First Boss: Pick a number (0/1)"


read firstboss

firstbossfight=$(( $RANDOM % 2 ))

if [[ $firstboss == $firstbossfight ]]; then
        echo "You vanquished the beast"
else
        echo "You Died"
        exit 1
fi

sleep 3

echo "Margit approaches, pick a number 0-9 (0-0)"

read secondboss

secondbossfight=$(( $RANDOM % 10 ))

if [[ $secondboss == $secondbossfight || $secondboss == swordoftruth ]]; then
        echo "YOU WON"
else
        echo "YOU DIED"
fi