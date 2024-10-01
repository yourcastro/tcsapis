import React, { useState } from "react";

import { Link, useNavigate } from "react-router-dom/dist";
import { ButtonDropdown, DropdownToggle, DropdownMenu, DropdownItem, Label, Dropdown } from 'reactstrap';


const ButtonDropDownComponent = ({ label, options, mappingOptions, ClassificationOption, classificationMainData }) => {
    console.log('label', label)

    let navigate = useNavigate();
    const [dropDownOpen, setDropDownOpen] = useState(false)
    const [subDropdown, setSubDropDown] = useState(false)
    const [ClassificationDropDown, setclassificationdropDown] = useState(false)
    const [clssificationDropDownMain, setclssificationDropDownMain] = useState(false)
    const [mappingSelect, setMappingSelect] = useState(null)

    // const navigateTo = () => history.push("/About");//eg.history.push('/login');
    const toggleDropDown = () => {
        setDropDownOpen(!dropDownOpen)
    }

    const toggleNested = () => {
        setSubDropDown(!subDropdown)
    }
    const toggleNestedClassification = () => {
        setclassificationdropDown(!ClassificationDropDown)
    }
    const toggleMainNestedClassification = () => {
        setclssificationDropDownMain(!clssificationDropDownMain)
    }
    const optionClick = (index, options) => {
        console.log('options', options)
        if (options === 'Home') {
            navigate("/LogIn");
        }
        if (options === 'Search Entities') {
            navigate("/SearchEntity");
        }
        if (options === 'Child/Parent Relationship') {
            navigate("/ChildParentReletionship");

        }
        if (options === 'Parent/Child Relationship') {
            navigate("/ParentChildReletionship");

        }
        if (options === 'Legal Entity Id Relationship') {
            navigate("/LeaglEntityReletionship");

        }
        if (options === 'Add Entity') {
            navigate("/AddEntity");

        }
        if (options === 'Deal Entity Relationship Management') {
            navigate("/DealEntityReletionship");

        }
        // if(options === 'Mappings'){
        //     setMappingSelect(true)
        //     setMappingSelect(options)
        //     // navigate("/Country")

        // }
        // if(options === 'Classification Data'){
        //     // setMappingSelect(true)
        //     navigate("/ClassificationData")

        // }
        if (options === 'Search Asset LGD Scorecards') {
            navigate("/LGDScoreCards");


        }
        if (options === 'Search Archived Asset LGD Scorecards') {
            navigate("/SearchArchivedAsset");


        }
        if (options === 'PD Scorecard Maintenance') {
            navigate("/PdScoreCardMaintanence");


        }
        if (options === 'Mortgage EL Rating Details') {
            navigate("/MortgageElRatingDetails");


        }
        if (options === 'Publics, Privates, Derivatives EL Ratings Detail') {
            navigate("/PublicPrivateElRatings");
        }
        if (options === 'Control Tables') {
            navigate("/ControlTable");
        }





    }
    const mappingOptionClick = (index, mapOption) => {

        if (mapOption === 'Country') {

            navigate("/Country")

        }
    }
    const classificationOptionClick = (index, ClassificationOption) => {
        console.log('ClassificationOption', ClassificationOption)

        if (ClassificationOption[0] === 'Classification') {

            navigate("/ClassificationData")

        }

    }
    const classficationMaindataClick=(index, CertificationOptionsData)=>{
        console.log('CertificationOptionsData',CertificationOptionsData)
        if (CertificationOptionsData === 'Bloomberg to GICS') {

            navigate("/BloomBergToGics")

        } 
        if (CertificationOptionsData === 'Lehmans To GICS') {

            navigate("/LehmansToGics")

        }  
        if (CertificationOptionsData === 'MSCI To GICS') {

            navigate("/MSCIToGics")

        } 
        if (CertificationOptionsData === 'unmapped GICS') {

            navigate("/UnmappedGics")

        } 
         
       

    }

    return (
        <div>

            <Dropdown
                isOpen={dropDownOpen}
                toggle={toggleDropDown}
                onMouseEnter={() => {
                    setDropDownOpen(true)
                }}
                onMouseLeave={() => {
                    setDropDownOpen(false)
                }}

            >
                <DropdownToggle className="DropDownToggle skinColor" >{label}</DropdownToggle>
                <DropdownMenu className="DropDownMenu">{options.map((options, index) =>
                    <DropdownItem className="DropDownItem" onClick={() => optionClick(index, options)} key={index}>{options}</DropdownItem>)}
                    {label === 'Entity Datastore' ?
                        <div>

                            <Dropdown onMouseEnter={toggleNested}
                                onMouseLeave={toggleNested}
                                isOpen={subDropdown}
                                toggle={toggleNested}
                                direction="right"
                            >
                                <DropdownToggle className="DropDownToggle skinColor" caret >Mapping </DropdownToggle>

                                <DropdownMenu className="DropDownMenu">{mappingOptions.map((mapOptions, index) =>
                                    <DropdownItem className="DropDownItem" onClick={() => mappingOptionClick(index, mapOptions)} key={index} >{mapOptions}</DropdownItem>)}
                                    <div>

                                        <Dropdown onMouseEnter={toggleNestedClassification}
                                            onMouseLeave={toggleNestedClassification}
                                            isOpen={ClassificationDropDown}
                                            toggle={toggleNestedClassification}
                                            direction="right"
                                        >
                                            <DropdownToggle className="DropDownToggle skinColor" caret >Classification Data</DropdownToggle>

                                            <DropdownMenu className="DropDownMenu">{ClassificationOption.map((CertificationOptions, index) =>
                                                <DropdownItem className="DropDownItem" onClick={() => classificationOptionClick(index, ClassificationOption)} key={index} >{CertificationOptions}</DropdownItem>)}
                                                <div>

                                                    <Dropdown onMouseEnter={toggleMainNestedClassification}
                                                        onMouseLeave={toggleMainNestedClassification}
                                                        isOpen={clssificationDropDownMain}
                                                        toggle={toggleMainNestedClassification}
                                                        direction="right"
                                                    >
                                                        <DropdownToggle className="DropDownToggle skinColor" caret >Classification Mapping</DropdownToggle>

                                                        <DropdownMenu className="DropDownMenu">{classificationMainData.map((CertificationOptionsData, index) =>
                                                            <DropdownItem className="DropDownItem" onClick={() => classficationMaindataClick(index, CertificationOptionsData)} key={index} >{CertificationOptionsData}</DropdownItem>)}
                                                        </DropdownMenu>
                                                    </Dropdown>
                                                </div>
                                            </DropdownMenu>
                                        </Dropdown>
                                    </div>
                                </DropdownMenu>
                            </Dropdown>
                        </div> : null}
                </DropdownMenu>

            </Dropdown>

            {/* <Header/>

            <h1>LohIn</h1>
            <Button onClick={checkOut} color="danger">DataShow</Button> */}

        </div>
    )
}
export default ButtonDropDownComponent;
