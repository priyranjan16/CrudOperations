import * as React from 'react';

export default class ParentClass extends React.Component<{}>{
    public render():React.ReactElement<{}>{
        return(
            <>
            <ChlidClass/>
            </>
        )
    }
}

class ChlidClass extends ParentClass{
    public render():React.ReactElement<{}>{
        return(
            <>
            <p>Hi, I am child class.</p>
            </>
        )
    }
}