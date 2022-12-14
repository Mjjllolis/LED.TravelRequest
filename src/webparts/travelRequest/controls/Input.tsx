import * as React from 'react';


const ReqInput = (props) => {
	//console.log(props.value);
	return (  
  <div className="form-group">
    <label className="form-label">{props.title}</label>
    
    <input
      className="form-control"
      id={props.name}
      name={props.name}
      type={props.inputType}
      value={props.value}
      onChange={props.handleChange}
      placeholder={props.placeholder} 
      {...props} />
  </div>
);
};

export default ReqInput;