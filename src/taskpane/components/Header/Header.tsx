import * as React from "react";
import "./Header.scss"

export interface HeaderProps {
  title: string;
  logo: string;
  message: string;
}

const Header = (props: HeaderProps) => {
  const { title, logo, message } = props;

  return (
    <section className="ms-welcome__header ms-bgColor-neutralLighter ms-u-fadeIn500 HeaderConatiner">
      <img width="300px" height="90" src={logo} alt={title} title={title} />
    </section>
  );
};
export { Header };
